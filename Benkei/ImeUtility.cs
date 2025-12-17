using System;
using System.Runtime.InteropServices;

namespace Benkei
{
    internal static class ImeUtility
    {
        private const int WmImeControl = 0x0283;
        private const int ImcSetopenstatus = 0x0006;
        private const int ImcGetopenstatus = 0x0005;
        private const int IMC_SETCONVERSIONMODE = 0x0002;
        private const int ImmGetconversionmode = 0x0001;
        private const int ImeCmodeNative = 0x0001;
        private const int ImeCmodeKatakana = 0x0002;
        private const int ImeCmodeFullshape = 0x0008;
        private const int ImeCmodeRoman = 0x0010;
        private const int CModeHiragana = ImeCmodeNative | ImeCmodeFullshape | ImeCmodeRoman;
        private const int CModeKatakana = ImeCmodeNative | ImeCmodeFullshape | ImeCmodeKatakana;
        private const int GcsCompstr = 0x0008;

        public static bool IsJapaneseInputActive()
        {
            return IsJapaneseInputActive1() || IsJapaneseInputActive2();
        }

        public static bool IsJapaneseInputActive1()
        {
            if (!TryGetFocusedWindow(out var foreground))
            {
                Logger.Log("[Interceptor] フォアグラウンドウィンドウの取得に失敗");
                return false;
            }

            var threadId = GetWindowThreadProcessId(foreground, out _);
            var layout = GetKeyboardLayout(threadId);
            var languageId = layout.ToInt64() & 0xFFFF;
            if (languageId != 0x0411)
            {
                Logger.Log("[Interceptor] 日本語入力以外のキーボードレイアウトがアクティブ");
                return false;
            }

            if (!TryGetDefaultContext(out var defaultContext))
            {
                Logger.Log("[Interceptor] デフォルトIMEウィンドウの取得に失敗");
                return false;
            }

            var result = SendMessage(defaultContext, WmImeControl, (IntPtr)ImcGetopenstatus, IntPtr.Zero);
            var isOpen = result.ToInt32() != 0;

            if (!isOpen)
            {
                Logger.Log("[Interceptor] IME==オフ");
                return false;
            }

            var conversionResult = SendMessage(defaultContext, WmImeControl, (IntPtr)ImmGetconversionmode, IntPtr.Zero);
            var conversion = conversionResult.ToInt32();

            var isNativeMode = (conversion & ImeCmodeNative) != 0;
            Logger.Log($"[Interceptor] IME変換モード: 0x{conversion:X}, ネイティブ: {isNativeMode}");
            return isNativeMode;
        }

        public static bool IsJapaneseInputActive2()
        {
            var foreground = GetForegroundWindow();
            if (foreground == IntPtr.Zero)
            {
                Logger.Log("[Interceptor] フォアグラウンドウィンドウの取得に失敗");
                return false;
            }

            var threadId = GetWindowThreadProcessId(foreground, out _);
            var layout = GetKeyboardLayout(threadId);
            var languageId = layout.ToInt64() & 0xFFFF;
            if (languageId != 0x0411)
            {
                Logger.Log("[Interceptor] 日本語入力以外のキーボードレイアウトがアクティブ");
                return false;
            }

            var context = ImmGetContext(foreground);
            if (context == IntPtr.Zero)
            {
                Logger.Log("[Interceptor] IMEコンテキストの取得に失敗");
                return false;
            }

            try
            {
                if (!ImmGetOpenStatus(context))
                {
                    Logger.Log("[Interceptor] IMEがオープンではありません");
                    return false;
                }

                const int ImeCmodeNative = 0x0001;
                if (ImmGetConversionStatus(context, out var conversion, out _))
                {
                    Logger.Log("[Interceptor] IMEの変換モードを確認");
                    return (conversion & ImeCmodeNative) != 0;
                }

                Logger.Log("[Interceptor] 日本語入力モード");
                return true;
            }
            finally
            {
                ImmReleaseContext(foreground, context);
            }
        }

        public static bool TryTurnOnHiragana()
        {
            if (!TryGetDefaultContext(out var defaultContext))
            {
                return false;
            }

            SendMessage(defaultContext, WmImeControl, (IntPtr)ImcSetopenstatus, (IntPtr)1);
            SendMessage(defaultContext, WmImeControl, (IntPtr)IMC_SETCONVERSIONMODE, (IntPtr)CModeHiragana);
            return true;
        }

        public static bool TryTurnOff()
        {
            if (!TryGetDefaultContext(out var defaultContext))
            {
                return false;
            }

            SendMessage(defaultContext, WmImeControl, (IntPtr)ImcSetopenstatus, (IntPtr)0);
            return true;
        }

        public static bool TryTurnOnKatakana()
        {
            if (!TryGetDefaultContext(out var defaultContext))
            {
                return false;
            }

            SendMessage(defaultContext, WmImeControl, (IntPtr)ImcSetopenstatus, (IntPtr)1);
            SendMessage(defaultContext, WmImeControl, (IntPtr)IMC_SETCONVERSIONMODE, (IntPtr)CModeKatakana);
            return true;
        }

        // 未変換文字が存在するかどうかを取得します。
        public static bool TryHasUnconvertedText()
        {
            if (!TryGetFocusedWindow(out var focused))
            {
                return false;
            }

            var inputContext = ImmGetContext(focused);
            if (inputContext == IntPtr.Zero)
            {
                return false;
            }

            try
            {
                var bytesNeeded = ImmGetCompositionString(inputContext, GcsCompstr, null, 0);
                if (bytesNeeded < 0) // エラー
                {
                    return false;
                }
                
                return bytesNeeded > 0;
            }
            finally
            {
                ImmReleaseContext(focused, inputContext);
            }
        }

        private static bool TryGetDefaultContext(out IntPtr defaultContext)
        {
            if (!TryGetFocusedWindow(out var focused))
            {
                defaultContext = IntPtr.Zero;
                return false;
            }

            defaultContext = ImmGetDefaultIMEWnd(focused);
            return defaultContext != IntPtr.Zero;
        }

        private static bool TryGetFocusedWindow(out IntPtr focused)
        {
            focused = IntPtr.Zero;

            var foreground = GetForegroundWindow();
            if (foreground == IntPtr.Zero)
            {
                return false;
            }

            var threadId = GetWindowThreadProcessId(foreground, out _);
            if (threadId == 0)
            {
                return false;
            }

            var info = new GuiThreadInfo
            {
                cbSize = (uint)Marshal.SizeOf<GuiThreadInfo>()
            };

            if (!GetGUIThreadInfo(threadId, ref info))
            {
                return false;
            }

            focused = info.hwndFocus != IntPtr.Zero ? info.hwndFocus : info.hwndActive;
            return focused != IntPtr.Zero;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct GuiThreadInfo
        {
            public uint cbSize;
            public uint flags;
            public IntPtr hwndActive;
            public IntPtr hwndFocus;
            public IntPtr hwndCapture;
            public IntPtr hwndMenuOwner;
            public IntPtr hwndMoveSize;
            public IntPtr hwndCaret;
            public RECT rcCaret;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int left;
            public int top;
            public int right;
            public int bottom;
        }

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll")]
        private static extern IntPtr GetKeyboardLayout(uint idThread);

        [DllImport("user32.dll")]
        private static extern bool GetGUIThreadInfo(uint idThread, ref GuiThreadInfo lpgui);

        [DllImport("imm32.dll")]
        private static extern IntPtr ImmGetDefaultIMEWnd(IntPtr hWnd);

        [DllImport("imm32.dll")]
        private static extern IntPtr ImmGetContext(IntPtr hWnd);

        [DllImport("imm32.dll")]
        private static extern bool ImmReleaseContext(IntPtr hWnd, IntPtr hIMC);

        [DllImport("imm32.dll", CharSet = CharSet.Unicode)]
        private static extern int ImmGetCompositionString(IntPtr hIMC, int dwIndex, byte[] lpBuf, int dwBufLen);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        [DllImport("imm32.dll")]
        private static extern bool ImmGetOpenStatus(IntPtr hIMC);

        [DllImport("imm32.dll")]
        private static extern bool ImmGetConversionStatus(IntPtr hIMC, out int conversion, out int sentence);
    }
}
