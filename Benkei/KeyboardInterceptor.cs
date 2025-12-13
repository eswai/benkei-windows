using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Benkei
{
    internal sealed class KeyboardInterceptor : IDisposable
    {
        private readonly NaginataEngine _engine;
        private readonly KeyActionExecutor _executor;
        private readonly AlphabetConfig _alphabetConfig;
        private readonly HashSet<int> _pressedPhysicalKeys = new HashSet<int>();
        private readonly LowLevelKeyboardProc _callback;
        private IntPtr _hookHandle = IntPtr.Zero;
        private bool _allowRepeat;
        private bool _isRepeating;
        private volatile bool _conversionEnabled = true;
        private int hjbuf = -1; // HJ,FG同時押しバッファ

        public KeyboardInterceptor(NaginataEngine engine, AlphabetConfig alphabetConfig)
        {
            _engine = engine ?? throw new ArgumentNullException(nameof(engine));
            _alphabetConfig = alphabetConfig ?? throw new ArgumentNullException(nameof(alphabetConfig));
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
        }

        public void Stop()
        {
            if (_hookHandle == IntPtr.Zero)
            {
                return;
            }

            UnhookWindowsHookEx(_hookHandle);
            _hookHandle = IntPtr.Zero;
            ResetStateInternal();
        }

        public void Dispose()
        {
            Stop();
        }

        public void SetConversionEnabled(bool enabled)
        {
            var previous = _conversionEnabled;
            _conversionEnabled = enabled;

            if (!enabled)
            {
                Console.WriteLine("[Interceptor] 入力変換 OFF");
                ResetStateInternal();
            }
            else if (!previous && enabled)
            {
                Console.WriteLine("[Interceptor] 入力変換 ON");
            }
        }

        public bool GetConversionEnabled() => _conversionEnabled;

        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
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

                if (!_conversionEnabled)
                {
                    return CallNextHookEx(_hookHandle, nCode, wParam, lParam);
                }

                var isJapaneseInputActive = ImeUtility.IsJapaneseInputActive();

                if (!isJapaneseInputActive && TryHandleImeOffKey(keyCode, isKeyDown, isKeyUp))
                {
                    return (IntPtr)1;
                }

                if (!_engine.IsNaginataKey(keyCode) || !isJapaneseInputActive)
                {
                    if (isKeyUp)
                    {
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
                        Console.WriteLine($"[Interceptor] リピートブロック: {keyCode}");
                        return (IntPtr)1;
                    }

                    Console.WriteLine($"[Interceptor] KeyDown: {keyCode}");
                    var actions = _engine.HandleKeyDown(keyCode);
                    Console.WriteLine($"[Interceptor] アクション数: {actions.Count}");
                    _executor.Execute(actions);
                    return (IntPtr)1;
                }

                if (isKeyUp)
                {
                    _pressedPhysicalKeys.Remove(keyCode);
                    _isRepeating = false;
                    _allowRepeat = false;
                    Console.WriteLine($"[Interceptor] KeyUp: {keyCode}");
                    var actions = _engine.HandleKeyUp(keyCode);
                    Console.WriteLine($"[Interceptor] アクション数: {actions.Count}");
                    _executor.Execute(actions);
                    return (IntPtr)1;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Interceptor] フックエラー: {ex.Message}");
                Console.WriteLine($"[Interceptor] スタックトレース: {ex.StackTrace}");
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
                            Console.WriteLine("[Interceptor] IME ON トグル");
                            IMEON();
                            hjbuf = -1;
                            return true;
                        }

                        if (IsImeOffCombo(hjbuf, keyCode))
                        {
                            Console.WriteLine("[Interceptor] IME OFF トグル");
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
                Console.WriteLine("[Interceptor] IME ON 失敗");
            }
            _engine.Reset();
        }

        private void IMEOFF()
        {
            if (!ImeUtility.TryTurnOff())
            {
                Console.WriteLine("[Interceptor] IME OFF 失敗");
            }
            _engine.Reset();
        }

        private void SetRepeatAllowed(bool allowed)
        {
            _allowRepeat = allowed;
        }

        private void ResetStateInternal()
        {
            _engine.Reset();
            _pressedPhysicalKeys.Clear();
            _isRepeating = false;
            _allowRepeat = false;
            _executor.ReleaseLatchedKeys();
            hjbuf = -1;
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
