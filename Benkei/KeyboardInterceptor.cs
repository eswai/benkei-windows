using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Benkei
{
    internal sealed class KeyboardInterceptor : IDisposable
    {
        private readonly NaginataEngine _engine;
        private readonly KeyActionExecutor _executor;
        private readonly HashSet<int> _pressedPhysicalKeys = new HashSet<int>();
        private readonly LowLevelKeyboardProc _callback;
        private IntPtr _hookHandle = IntPtr.Zero;
        private bool _allowRepeat;
        private bool _isRepeating;

        public KeyboardInterceptor(NaginataEngine engine)
        {
            _engine = engine ?? throw new ArgumentNullException(nameof(engine));
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

                if (!_engine.IsNaginataKey(keyCode) || !IsJapaneseInputActive())
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
        }

        private bool IsJapaneseInputActive()
        {
            var foreground = GetForegroundWindow();
            if (foreground == IntPtr.Zero)
            {
                Console.WriteLine("[Interceptor] フォアグラウンドウィンドウの取得に失敗");
                return false;
            }

            var threadId = GetWindowThreadProcessId(foreground, out _);
            var layout = GetKeyboardLayout(threadId);
            var languageId = layout.ToInt64() & 0xFFFF;
            if (languageId != 0x0411)
            {
                Console.WriteLine("[Interceptor] 日本語入力以外のキーボードレイアウトがアクティブ");
                return false;
            }

            var defaultContext = ImmGetDefaultIMEWnd(foreground);
            if (defaultContext == IntPtr.Zero)
            {
                Console.WriteLine("[Interceptor] デフォルトIMEウィンドウの取得に失敗");
                return false;
            }

            // IMEの状態を取得
            var result = SendMessage(defaultContext, WmImeControl, new IntPtr(ImcGetopenstatus), IntPtr.Zero);
            var isOpen = result.ToInt32() != 0;
            
            if (!isOpen)
            {
                Console.WriteLine("[Interceptor] IMEがオフ（英数モード）");
                return false;
            }

            // 変換モードを取得
            var conversionResult = SendMessage(defaultContext, WmImeControl, new IntPtr(ImmGetconversionmode), IntPtr.Zero);
            var conversion = conversionResult.ToInt32();
            
            const int ImeCmodeNative = 0x0001;
            var isNativeMode = (conversion & ImeCmodeNative) != 0;
            Console.WriteLine($"[Interceptor] IME変換モード: 0x{conversion:X}, ネイティブ: {isNativeMode}");
            return isNativeMode;
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
        private const int WmImeControl = 0x0283;
        private const int ImmGetconversionmode = 0x0001;
        private const int ImcGetopenstatus = 0x0005;
        private static readonly IntPtr BenkeiMarker = new IntPtr(0x42454E4B); // "BENK"

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        [DllImport("user32.dll")]
        private static extern IntPtr GetKeyboardLayout(uint idThread);

        [DllImport("imm32.dll")]
        private static extern IntPtr ImmGetDefaultIMEWnd(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);
    }
}
