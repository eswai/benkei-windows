using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace Benkei
{
    internal struct KeyEventData
    {
        public int KeyCode;
        public bool IsKeyDown;
        public bool IsKeyUp;
        public bool UseImeOnEngine;
    }

    internal sealed class KeyboardInterceptor : IDisposable
    {
        private readonly NaginataEngine _imeOnEngine;
        private readonly NaginataEngine _imeOffEngine;
        private readonly KeyActionExecutor _executor;
        private readonly Action<bool> _conversionStateChanged;
        private readonly HashSet<int> _pressedPhysicalKeys = new HashSet<int>();
        private readonly Dictionary<int, bool> _keyEngineBindings = new Dictionary<int, bool>();
        private readonly LowLevelKeyboardProc _callback;
        private IntPtr _hookHandle = IntPtr.Zero;
        private bool _allowRepeat;
        private bool _isRepeating;
        private volatile bool _conversionEnabled = true;
        private bool _ctrlPressed;
        private bool _shiftPressed;
        private bool _altPressed;
        private readonly ConcurrentQueue<KeyEventData> _keyEventQueue = new ConcurrentQueue<KeyEventData>();
        private readonly AutoResetEvent _keyEventSignal = new AutoResetEvent(false);
        private Thread _workerThread;
        private volatile bool _workerRunning;

        public KeyboardInterceptor(NaginataEngine imeOnEngine, NaginataEngine imeOffEngine, Action<bool> conversionStateChanged = null)
        {
            _imeOnEngine = imeOnEngine ?? throw new ArgumentNullException(nameof(imeOnEngine));
            _imeOffEngine = imeOffEngine ?? throw new ArgumentNullException(nameof(imeOffEngine));
            _conversionStateChanged = conversionStateChanged;
            _callback = HookCallback;
            _executor = new KeyActionExecutor(SetRepeatAllowed, ResetStateInternal);
        }

        public void Start()
        {
            if (_hookHandle != IntPtr.Zero)
            {
                return;
            }

            IntPtr moduleHandle;
            using (var process = Process.GetCurrentProcess())
            {
                using (var module = process.MainModule)
                {
                    var moduleName = module != null ? module.ModuleName : null;
                    moduleHandle = GetModuleHandle(moduleName);
                }
            }

            if (moduleHandle == IntPtr.Zero)
            {
                moduleHandle = GetModuleHandle(null);
            }

            _hookHandle = SetWindowsHookEx(WhKeyboardLl, _callback, moduleHandle, 0);
            if (_hookHandle == IntPtr.Zero)
            {
                throw new InvalidOperationException("Failed to install keyboard hook.");
            }

            // ワーカースレッドの起動
            _workerRunning = true;
            _workerThread = new Thread(ProcessKeyEventQueue)
            {
                IsBackground = true,
                Name = "KeyEventProcessor"
            };
            _workerThread.Start();
            Logger.Log("[Interceptor] ワーカースレッド起動");
        }

        public void Stop()
        {
            if (_hookHandle == IntPtr.Zero)
            {
                return;
            }

            UnhookWindowsHookEx(_hookHandle);
            _hookHandle = IntPtr.Zero;

            // ワーカースレッドの停止
            _workerRunning = false;
            _keyEventSignal.Set();
            if (_workerThread != null && _workerThread.IsAlive)
            {
                if (!_workerThread.Join(1000))
                {
                    Logger.Log("[Interceptor] ワーカースレッド停止タイムアウト");
                }
            }
            Logger.Log("[Interceptor] ワーカースレッド停止");

            ResetStateInternal();
        }

        public void Dispose()
        {
            Stop();
            _keyEventSignal?.Dispose();
        }

        public void SetConversionEnabled(bool enabled)
        {
            var previous = _conversionEnabled;
            _conversionEnabled = enabled;

            if (!enabled)
            {
                Logger.Log("[Interceptor] 入力変換 OFF");
                ResetStateInternal();
            }
            else if (!previous && enabled)
            {
                Logger.Log("[Interceptor] 入力変換 ON");
            }

            _conversionStateChanged?.Invoke(enabled);
        }

        public bool GetConversionEnabled() => _conversionEnabled;

        // フック内ログは遅延の主要因になりやすいので、通常はOFF推奨
        private const bool HookVerboseLog = false;

        private static void HookLog(string message)
        {
            if (HookVerboseLog)
            {
                Logger.Log(message);
            }
        }

        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (!_conversionEnabled)
            {
                return CallNextHookEx(_hookHandle, nCode, wParam, lParam);
            }

            if (nCode < 0)
            {
                return CallNextHookEx(_hookHandle, nCode, wParam, lParam);
            }

            try
            {
                var message = wParam.ToInt32();
                var isKeyDown = message == WmKeydown || message == WmSyskeydown;
                var isKeyUp = message == WmKeyup || message == WmSyskeyup;

                if (!isKeyDown && !isKeyUp)
                {
                    return CallNextHookEx(_hookHandle, nCode, wParam, lParam);
                }

                var hookData = Marshal.PtrToStructure<KbdLlHookStruct>(lParam);
                var keyCode = hookData.vkCode;

                // Skip events sent by ourselves
                if (hookData.dwExtraInfo == BenkeiMarker)
                {
                    return CallNextHookEx(_hookHandle, nCode, wParam, lParam);
                }

                if (TryUpdateModifierState(keyCode, isKeyDown, isKeyUp))
                {
                    return CallNextHookEx(_hookHandle, nCode, wParam, lParam);
                }

                if (TryHandleConversionToggle(keyCode, isKeyDown))
                {
                    return (IntPtr)1;
                }

                var isJapaneseInputActive = ImeUtility.IsJapaneseInputActive();
                if (isJapaneseInputActive)
                {
                    var isNaginataKey = _imeOnEngine.IsNaginataKey(keyCode);
                    if (!isNaginataKey)
                    {
                        return CallNextHookEx(_hookHandle, nCode, wParam, lParam);
                    }
                } else {
                    var isConversionKey = _imeOffEngine.IsConversionKey(keyCode);
                    if (!isConversionKey)
                    {
                        return CallNextHookEx(_hookHandle, nCode, wParam, lParam);
                    }
                }

                bool useImeOnEngine;
                if (_keyEngineBindings.TryGetValue(keyCode, out var existingBinding))
                {
                    useImeOnEngine = existingBinding;
                }
                else
                {
                    if (!isKeyDown)
                    {
                        return CallNextHookEx(_hookHandle, nCode, wParam, lParam);
                    }

                    useImeOnEngine = isJapaneseInputActive;
                    _keyEngineBindings[keyCode] = useImeOnEngine;
                }

                var targetEngine = useImeOnEngine ? _imeOnEngine : _imeOffEngine;
                if (targetEngine == null)
                {
                    _keyEngineBindings.Remove(keyCode);
                    return CallNextHookEx(_hookHandle, nCode, wParam, lParam);
                }

                if (isKeyDown)
                {
                    if (!_pressedPhysicalKeys.Add(keyCode))
                    {
                        _isRepeating = true;
                    }
                    else
                    {
                        _isRepeating = false;
                    }

                    if (_isRepeating && !_allowRepeat)
                    {
                        HookLog($"[Interceptor] リピートブロック: {keyCode}");
                        return (IntPtr)1;
                    }

                    HookLog($"[Interceptor] KeyDown: {keyCode}");
                    _keyEventQueue.Enqueue(new KeyEventData { KeyCode = keyCode, IsKeyDown = true, IsKeyUp = false, UseImeOnEngine = useImeOnEngine });
                    _keyEventSignal.Set();
                    return (IntPtr)1;
                }

                if (isKeyUp)
                {
                    if (!_pressedPhysicalKeys.Contains(keyCode))
                    {
                        _keyEngineBindings.Remove(keyCode);
                        return CallNextHookEx(_hookHandle, nCode, wParam, lParam);
                    }

                    _pressedPhysicalKeys.Remove(keyCode);

                    if (_isRepeating)
                    {
                        while (_keyEventQueue.TryDequeue(out _)) { }
                        HookLog($"[Interceptor] リピート解除 & キュークリア: {keyCode}");
                    }

                    _isRepeating = false;
                    _allowRepeat = false;
                    HookLog($"[Interceptor] KeyUp: {keyCode}");
                    _keyEngineBindings.Remove(keyCode);
                    _keyEventQueue.Enqueue(new KeyEventData { KeyCode = keyCode, IsKeyDown = false, IsKeyUp = true, UseImeOnEngine = useImeOnEngine });

                    _keyEventSignal.Set();
                    return (IntPtr)1;
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"[Interceptor] フックエラー: {ex.Message}");
                Logger.Log($"[Interceptor] スタックトレース: {ex.StackTrace}");
                ResetStateInternal();
            }

            return CallNextHookEx(_hookHandle, nCode, wParam, lParam);
        }

        private void SetRepeatAllowed(bool allowed)
        {
            _allowRepeat = allowed;
        }

        private bool TryHandleConversionToggle(int keyCode, bool isKeyDown)
        {
            if (!isKeyDown)
            {
                return false;
            }

            if (!_ctrlPressed || !_shiftPressed)
            {
                return false;
            }

            if (keyCode == (int)Keys.D0)
            {
                SetConversionEnabled(false);
                return true;
            }

            if (keyCode == (int)Keys.D1)
            {
                SetConversionEnabled(true);
                return true;
            }

            return false;
        }

        private void ResetStateInternal()
        {
            _imeOnEngine.Reset();
            _imeOffEngine.Reset();
            _pressedPhysicalKeys.Clear();
            _keyEngineBindings.Clear();
            _isRepeating = false;
            _allowRepeat = false;
            _executor.ReleaseLatchedKeys();
            _ctrlPressed = false;
            _shiftPressed = false;
            _altPressed = false;
        }

        private static bool IsControlKey(int keyCode)
        {
            return keyCode == (int)Keys.ControlKey || keyCode == (int)Keys.RControlKey || keyCode == (int)Keys.LControlKey;
        }

        private static bool IsShiftKey(int keyCode)
        {
            return keyCode == (int)Keys.ShiftKey || keyCode == (int)Keys.RShiftKey || keyCode == (int)Keys.LShiftKey;
        }

        private static bool IsAltKey(int keyCode)
        {
            return keyCode == (int)Keys.Menu || keyCode == (int)Keys.RMenu;
        }

        private bool TryUpdateModifierState(int keyCode, bool isKeyDown, bool isKeyUp)
        {
            if (IsControlKey(keyCode))
            {
                _ctrlPressed = isKeyDown ? true : isKeyUp ? false : _ctrlPressed;
                return true;
            }

            if (IsShiftKey(keyCode))
            {
                _shiftPressed = isKeyDown ? true : isKeyUp ? false : _shiftPressed;
                return true;
            }

            if (IsAltKey(keyCode))
            {
                _altPressed = isKeyDown ? true : isKeyUp ? false : _altPressed;
                return true;
            }

            return false;
        }

        private void ProcessKeyEventQueue()
        {
            while (_workerRunning)
            {
                _keyEventSignal.WaitOne();

                while (_keyEventQueue.TryDequeue(out var eventData))
                {
                    try
                    {
                        var engine = eventData.UseImeOnEngine ? _imeOnEngine : _imeOffEngine;
                        if (engine == null)
                        {
                            continue;
                        }

                        List<NaginataAction> actions;
                        if (eventData.IsKeyDown)
                        {
                            actions = engine.HandleKeyDown(eventData.KeyCode);
                        }
                        else if (eventData.IsKeyUp)
                        {
                            actions = engine.HandleKeyUp(eventData.KeyCode);
                        }
                        else
                        {
                            continue;
                        }

                        HookLog($"[Worker] アクション数: {actions.Count}");
                        _executor.Execute(actions);
                    }
                    catch (Exception ex)
                    {
                        Logger.Log($"[Worker] キー処理エラー: {ex.Message}");
                    }
                }
            }
        }

        private void RestoreKanaModeIfNeeded()
        {
            if (_conversionEnabled)
            {
                ImeUtility.TryTurnOnHiragana();
            }
        }

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        private struct KbdLlHookStruct
        {
            public int vkCode;
            public int scanCode;
            public int flags;
            public int time;
            public IntPtr dwExtraInfo;
        }

        private const int WhKeyboardLl = 13;
        private const int WmKeydown = 0x0100;
        private const int WmKeyup = 0x0101;
        private const int WmSyskeydown = 0x0104;
        private const int WmSyskeyup = 0x0105;
        private static readonly IntPtr BenkeiMarker = new IntPtr(0x42454E4B); // "BENK"

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

    }
}
