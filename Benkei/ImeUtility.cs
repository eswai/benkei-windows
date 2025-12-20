using System;
using System.Runtime.InteropServices;

namespace Benkei
{
    internal static class ImeUtility
    {
        // 0: 未確定(未選択), 1: IsJapaneseInputActiveLegacyMicrosoftIME, 2: IsJapaneseInputActiveCurrentMicrosoftIME
        // いったん true を返した方式を優先して以後の判定コストを下げる。
        private static int _preferredJapaneseInputActiveChecker;

        private const int WmImeControl = 0x0283;
        private const int ImcSetopenstatus = 0x0006;
        private const int ImcGetopenstatus = 0x0005;

        // 追加: WM_IME_CONTROL で合成文字列の長さを取得（IMM 系 IME 向け）
        private const int ImcGetcompositionstring = 0x0001;

        private const int IMC_SETCONVERSIONMODE = 0x0002;
        private const int ImmGetconversionmode = 0x0001;
        private const int ImeCmodeNative = 0x0001;
        private const int ImeCmodeKatakana = 0x0002;
        private const int ImeCmodeFullshape = 0x0008;
        private const int ImeCmodeRoman = 0x0010;
        private const int CModeHiragana = ImeCmodeNative | ImeCmodeFullshape | ImeCmodeRoman;
        private const int CModeKatakana = ImeCmodeNative | ImeCmodeFullshape | ImeCmodeKatakana;
        private const int GcsCompstr = 0x0008;

        // 追加: 旧 Microsoft IME だと GCS_COMPSTR が 0 のことがあるため併用
        private const int GcsCompreadstr = 0x0001;

        public static bool IsJapaneseInputActive()
        {
            // 最初に true になった方を記憶し、基本は片方だけ呼ぶ。
            // ただし、優先方式が false の場合のみもう片方を試して切り替える（環境差/一時失敗に強くする）。
            var preferred = System.Threading.Volatile.Read(ref _preferredJapaneseInputActiveChecker);
            if (preferred == 1)
            {
                if (IsJapaneseInputActiveLegacyMicrosoftIME())
                {
                    return true;
                }
                return false;
            }

            if (preferred == 2)
            {
                if (IsJapaneseInputActiveCurrentMicrosoftIME())
                {
                    return true;
                }
                return false;
            }

            // 未選択: 1 → 2 の順で試して、最初に true になった方を記憶する
            if (IsJapaneseInputActiveLegacyMicrosoftIME())
            {
                System.Threading.Volatile.Write(ref _preferredJapaneseInputActiveChecker, 1);
                Logger.Log("[Interceptor] 日本語入力判定: チェッカーを IsJapaneseInputActiveLegacyMicrosoftIME に固定");
                return true;
            }

            if (IsJapaneseInputActiveCurrentMicrosoftIME())
            {
                System.Threading.Volatile.Write(ref _preferredJapaneseInputActiveChecker, 2);
                Logger.Log("[Interceptor] 日本語入力判定: チェッカーを IsJapaneseInputActiveCurrentMicrosoftIME に固定");
                return true;
            }

            return false;
        }

        public static bool IsJapaneseInputActiveCurrentMicrosoftIME()
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

        public static bool IsJapaneseInputActiveLegacyMicrosoftIME()
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
            // まず通常の方法で確認
            var preferred = System.Threading.Volatile.Read(ref _preferredJapaneseInputActiveChecker);
            if (preferred == 1)
            {
                return TryHasUnconvertedTextLegacyMicrosoftIme();
            }
            if (preferred == 2)
            {
                return TryHasUnconvertedTextCurrentMicrosoftIME();
            }

            if (TryHasUnconvertedTextLegacyMicrosoftIme())
            {
                System.Threading.Volatile.Write(ref _preferredJapaneseInputActiveChecker, 1);
                return true;
            }
            if (TryHasUnconvertedTextCurrentMicrosoftIME())
            {
                System.Threading.Volatile.Write(ref _preferredJapaneseInputActiveChecker, 2);
                return true;
            }

            return false;
        }

        public static bool TryHasUnconvertedTextCurrentMicrosoftIME()
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

        /// <summary>
        /// 旧「Microsoft IME (以前のバージョン)」向け: 未確定(未変換)の合成文字が存在するか。
        /// WM_IME_CONTROL(IMC_GETCOMPOSITIONSTRING) 経由で長さを取得します。
        /// </summary>
        public static bool TryHasUnconvertedTextLegacyMicrosoftIme()
        {
            if (!TryGetFocusedWindow(out var focused))
            {
                return false;
            }

            var imeWnd = ImmGetDefaultIMEWnd(focused);
            if (imeWnd == IntPtr.Zero)
            {
                return false;
            }

            // 旧 IME は COMPSTR が取れないケースがあるため COMPREADSTR も確認
            var len1 = SendMessage(imeWnd, WmImeControl, (IntPtr)ImcGetcompositionstring, (IntPtr)GcsCompstr).ToInt32();
            if (len1 > 0) return true;

            var len2 = SendMessage(imeWnd, WmImeControl, (IntPtr)ImcGetcompositionstring, (IntPtr)GcsCompreadstr).ToInt32();
            return len2 > 0;
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
