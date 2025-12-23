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
    }

    internal sealed class KeyboardInterceptor : IDisposable
    {
        private readonly NaginataEngine _engine;
        private readonly KeyActionExecutor _executor;
        private readonly AlphabetConfig _alphabetConfig;
        private readonly Action<bool> _conversionStateChanged;
        private readonly HashSet<int> _pressedPhysicalKeys = new HashSet<int>();
        private readonly LowLevelKeyboardProc _callback;
        private IntPtr _hookHandle = IntPtr.Zero;
        private bool _allowRepeat;
        private bool _isRepeating;
        private volatile bool _conversionEnabled = true;
        private int hjbuf = -1; // HJ,FG同時押しバッファ
        private bool _ctrlPressed;
        private bool _shiftPressed;
        private bool _altPressed;
        private bool _windowsPressed;
        private readonly ConcurrentQueue<KeyEventData> _keyEventQueue = new ConcurrentQueue<KeyEventData>();
        private readonly AutoResetEvent _keyEventSignal = new AutoResetEvent(false);
        private Thread _workerThread;
        private volatile bool _workerRunning;

        public KeyboardInterceptor(NaginataEngine engine, AlphabetConfig alphabetConfig, Action<bool> conversionStateChanged = null)
        {
            _engine = engine ?? throw new ArgumentNullException(nameof(engine));
            _alphabetConfig = alphabetConfig ?? throw new ArgumentNullException(nameof(alphabetConfig));
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

                // ここから「IME 状態が必要か？」を先に判定して、
                // 不要なキーは IME 状態問い合わせなしで即パスする（軽量化）
                var isNaginataKey = _engine.IsNaginataKey(keyCode);
                var isImeToggleCandidate = IsImeToggleCandidate(keyCode);

                var hasAlphabetMapping = _alphabetConfig != null && _alphabetConfig.TryGetRemappedKey(keyCode, out _);
                var anyModifierPressed = _ctrlPressed || _shiftPressed || _altPressed || _windowsPressed;

                var needsImeState =
                    isNaginataKey ||                      // Naginata処理の可否に必要
                    isImeToggleCandidate ||               // H/J/F/G の処理に必要
                    (hasAlphabetMapping && anyModifierPressed); // 修飾キー＋リマップ時に必要

                if (!needsImeState)
                {
                    // ここに来るキーはBenkei側で何もしないので即パス
                    return CallNextHookEx(_hookHandle, nCode, wParam, lParam);
                }

                var isJapaneseInputActive = ImeUtility.IsJapaneseInputActive();

                if (isJapaneseInputActive && anyModifierPressed && TryHandleAlphabetRemap(keyCode, isKeyDown, isKeyUp))
                {
                    return (IntPtr)1;
                }

                if (!isJapaneseInputActive && TryHandleImeOffKey(keyCode, isKeyDown, isKeyUp))
                {
                    return (IntPtr)1;
                }

                if (!isNaginataKey || !isJapaneseInputActive)
                {
                    if (isKeyUp)
                    {
                        // 記号入力で英字モードになっているときにキーアップすると、キーが押されっぱなしになる
                        if (_pressedPhysicalKeys.Contains(keyCode))
                        {
                            _keyEventQueue.Enqueue(new KeyEventData { KeyCode = keyCode, IsKeyDown = false, IsKeyUp = true });
                            _keyEventSignal.Set();
                        }
                        _pressedPhysicalKeys.Remove(keyCode);
                        _isRepeating = false;
                        _allowRepeat = false;

                    }

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
                    _keyEventQueue.Enqueue(new KeyEventData { KeyCode = keyCode, IsKeyDown = true, IsKeyUp = false });
                    _keyEventSignal.Set();
                    return (IntPtr)1;
                }

                if (isKeyUp)
                {
                    if (!_pressedPhysicalKeys.Contains(keyCode))
                    {
                        // 押されていないキーの離鍵はOSに渡す（IME OFF中に押されたキーの可能性があるため）
                        return CallNextHookEx(_hookHandle, nCode, wParam, lParam);
                    }

                    _pressedPhysicalKeys.Remove(keyCode);
                    
                    if (_isRepeating)
                    {
                        // リピート中にキーを離した場合、キューを空にする
                        while (_keyEventQueue.TryDequeue(out _)) { }
                        HookLog($"[Interceptor] リピート解除 & キュークリア: {keyCode}");
                    }

                    _isRepeating = false;
                    _allowRepeat = false;
                    HookLog($"[Interceptor] KeyUp: {keyCode}");
                    _keyEventQueue.Enqueue(new KeyEventData { KeyCode = keyCode, IsKeyDown = false, IsKeyUp = true });

                    // if (_pressedPhysicalKeys.Count == 0)
                    // {
                    //     _keyEventQueue.Enqueue(new KeyEventData { KeyCode = 0xffff, IsKeyDown = true, IsKeyUp = false });
                    // }

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

        private bool TryHandleImeOffKey(int keyCode, bool isKeyDown, bool isKeyUp)
        {
            if (isKeyDown)
            {
                if (hjbuf == -1 && IsImeToggleCandidate(keyCode))
                {
                    hjbuf = keyCode;
                    return true;
                }

                if (hjbuf > -1)
                {
                    if (IsImeToggleCandidate(keyCode))
                    {
                        if (IsImeOnCombo(hjbuf, keyCode))
                        {
                            Logger.Log("[Interceptor] IME ON トグル");
                            IMEON();
                            hjbuf = -1;
                            return true;
                        }

                        if (IsImeOffCombo(hjbuf, keyCode))
                        {
                            Logger.Log("[Interceptor] IME OFF トグル");
                            IMEOFF();
                            hjbuf = -1;
                            return true;
                        }

                        SendRemappedTap(hjbuf);
                        SendRemappedTap(keyCode);
                        hjbuf = -1;
                        return true;
                    }

                    SendRemappedTap(hjbuf);
                    hjbuf = -1;
                }
            }
            else if (isKeyUp)
            {
                if (hjbuf > -1 && hjbuf == keyCode)
                {
                    SendRemappedTap(hjbuf);
                    hjbuf = -1;
                    return true;
                }
            }

            return TryHandleAlphabetRemap(keyCode, isKeyDown, isKeyUp);
        }

        private bool TryHandleAlphabetRemap(int keyCode, bool isKeyDown, bool isKeyUp)
        {
            if (_alphabetConfig == null)
            {
                return false;
            }

            var hasMapping = _alphabetConfig.TryGetRemappedKey(keyCode, out _);
            if (!hasMapping)
            {
                return false;
            }

            if (isKeyDown)
            {
                SendRemappedTap(keyCode);
                return true;
            }

            if (isKeyUp)
            {
                return true;
            }

            return false;
        }

        private void SendRemappedTap(int keyCode)
        {
            var targetKey = ResolveRemappedKey(keyCode);
            _executor.TapKey((ushort)targetKey);
        }

        private int ResolveRemappedKey(int keyCode)
        {
            if (_alphabetConfig != null && _alphabetConfig.TryGetRemappedKey(keyCode, out var mapped))
            {
                return mapped;
            }

            return keyCode;
        }

        private static bool IsImeToggleCandidate(int keyCode)
        {
            return keyCode == (int)Keys.H || keyCode == (int)Keys.J || keyCode == (int)Keys.F || keyCode == (int)Keys.G;
        }

        private static bool IsImeOnCombo(int first, int second)
        {
            return (first == (int)Keys.H && second == (int)Keys.J) || (first == (int)Keys.J && second == (int)Keys.H);
        }

        private static bool IsImeOffCombo(int first, int second)
        {
            return (first == (int)Keys.F && second == (int)Keys.G) || (first == (int)Keys.G && second == (int)Keys.F);
        }

        private void IMEON()
        {
            if (!ImeUtility.TryTurnOnHiragana())
            {
                Logger.Log("[Interceptor] IME ON 失敗");
            }
            _engine.Reset();
        }

        private void IMEOFF()
        {
            if (!ImeUtility.TryTurnOff())
            {
                Logger.Log("[Interceptor] IME OFF 失敗");
            }
            _engine.Reset();
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
            _engine.Reset();
            _pressedPhysicalKeys.Clear();
            _isRepeating = false;
            _allowRepeat = false;
            _executor.ReleaseLatchedKeys();
            hjbuf = -1;
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

        private static bool IsWindowsKey(int keyCode)
        {
            return keyCode == (int)Keys.RWin || keyCode == (int)Keys.LWin;
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

            if (IsWindowsKey(keyCode))
            {
                _windowsPressed = isKeyDown ? true : isKeyUp ? false : _windowsPressed;
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
                        List<NaginataAction> actions;
                        // if (eventData.KeyCode == 0xffff)
                        // {
                        //     _engine.Reset();
                        //     continue;
                        // }
                        if (eventData.IsKeyDown)
                        {
                            actions = _engine.HandleKeyDown(eventData.KeyCode);
                        }
                        else if (eventData.IsKeyUp)
                        {
                            actions = _engine.HandleKeyUp(eventData.KeyCode);
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
