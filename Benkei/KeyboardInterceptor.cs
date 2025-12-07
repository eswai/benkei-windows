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
        private readonly HashSet<int> _pressedPhysicalKeys = new HashSet<int>();
        private readonly LowLevelKeyboardProc _callback;
        private IntPtr _hookHandle = IntPtr.Zero;
        private bool _allowRepeat;
        private bool _isRepeating;
        private volatile bool _conversionEnabled = true;
        private int hjbuf = -1; // HJ同時押しバッファ

        const int IMC_SETCONVERSIONMODE = 2;
        const int IME_CMODE_NATIVE    =  1;
        const int IME_CMODE_KATAKANA  =  2;
        const int IME_CMODE_FULLSHAPE =  8;
        const int IME_CMODE_ROMAN     = 16;
        const int CMode_Hiragana    = IME_CMODE_ROMAN | IME_CMODE_FULLSHAPE | IME_CMODE_NATIVE;

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

                if (IsJapaneseInputActive() == false) {
                // IsJapaneseInputActive == falseのとき
                // HとJを同時に押すと、IMEをONにする
                // kana_on同時押しの処理（マッピング後のキーコードで判定）
                // if type == .keyDown {
                //     if hjbuf == -1 {
                //         if kanaOnKeys.count >= 2 && (targetKeyCode == kanaOnKeys[0] || targetKeyCode == kanaOnKeys[1]) {
                //             hjbuf = originalKeyCode; // 元のキーコードを保存
                //             return nil;
                //         } else {
                //             // マッピングされたキーを送信
                //             postKeyEvent(keyCode: targetKeyCode, keyDown: true)
                //             return nil
                //         }
                    if (isKeyDown) {
                        if (hjbuf == -1)
                        {
                            if (keyCode == (int)Keys.H || keyCode == (int)Keys.J)
                            {
                                hjbuf = keyCode; // 元のキーコードを保存
                                return (IntPtr)1;
                            } else {
                                _executor.PressKey((ushort)hjbuf);
                                return CallNextHookEx(_hookHandle, nCode, wParam, lParam);
                            }
                //     } else {
                //         let hjbufMapped = abcMapping[hjbuf] ?? hjbuf
                //         if hjbufMapped + targetKeyCode == kanaOnKeys[0] + kanaOnKeys[1] {
                //             sendJISKanaKey()
                //             hjbuf = -1
                //             return nil
                //         } else {
                //             // バッファのキーとマッピングされたキーを両方送信
                //             let hjbufTargetKeyCode = abcMapping[hjbuf] ?? hjbuf
                //             postKeyEvent(keyCode: hjbufTargetKeyCode, keyDown: true)
                //             postKeyEvent(keyCode: hjbufTargetKeyCode, keyDown: false)
                //             postKeyEvent(keyCode: targetKeyCode, keyDown: true)
                //             pressedKeys.remove(hjbuf)
                //             hjbuf = -1
                //             return nil
                //         }
                //     }
                        }
                        else
                        {
                            if (hjbuf + keyCode == (int)Keys.H + (int)Keys.J)
                            {
                                Console.WriteLine("[Interceptor] IME ON トグル");
                                IMEON();
                                hjbuf = -1;
                                return (IntPtr)1;
                            } else {
                                _executor.TapKey((ushort)hjbuf);
                                _executor.PressKey((ushort)keyCode);
                                _pressedPhysicalKeys.Remove(hjbuf);
                                hjbuf = -1;
                                return (IntPtr)1;
                            }
                        }
                // } else if type == .keyUp {
                //     if hjbuf > -1 && hjbuf == originalKeyCode {
                //         let hjbufTargetKeyCode = abcMapping[hjbuf] ?? hjbuf
                //         postKeyEvent(keyCode: hjbufTargetKeyCode, keyDown: true)
                //         postKeyEvent(keyCode: hjbufTargetKeyCode, keyDown: false)
                //         pressedKeys.remove(hjbuf)
                //         hjbuf = -1
                //         return nil
                //     } else {
                //         // マッピングされたキーのキーアップを送信
                //         postKeyEvent(keyCode: targetKeyCode, keyDown: false)
                //         return nil
                //     }
                // }
                    }
                    else
                    {
                        if (hjbuf > -1 && hjbuf == keyCode)
                        {
                            _executor.TapKey((ushort)hjbuf);
                            _pressedPhysicalKeys.Remove(hjbuf);
                            hjbuf = -1;
                            return (IntPtr)1;
                        }
                        else
                        {
                            // マッピングされたキーのキーアップを送信
                            _executor.ReleaseKey((ushort)keyCode);
                            return (IntPtr)1;
                        }
                    }
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

        private void IMEON() 
        {
            var foreground = GetForegroundWindow();
            var defaultContext = ImmGetDefaultIMEWnd(foreground);
            if (defaultContext != IntPtr.Zero)
            {
                const int ImcSetopenstatus = 0x0006;
                SendMessage(defaultContext, WmImeControl, new IntPtr(ImcSetopenstatus), new IntPtr(1));
                SendMessage(defaultContext, WmImeControl, (IntPtr)IMC_SETCONVERSIONMODE, (IntPtr)CMode_Hiragana);
            }
            _engine.Reset();
        }

        private void IMEOFF() 
        {
            var foreground = GetForegroundWindow();
            var defaultContext = ImmGetDefaultIMEWnd(foreground);
            if (defaultContext != IntPtr.Zero)
            {
                const int ImcSetopenstatus = 0x0006;
                SendMessage(defaultContext, WmImeControl, new IntPtr(ImcSetopenstatus), new IntPtr(0));
                // SendMessage(defaultContext, WmImeControl, (IntPtr)IMC_SETCONVERSIONMODE, (IntPtr)IME_CMODE_NATIVE);
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
                Console.WriteLine("[Interceptor] IME==オフ");
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
