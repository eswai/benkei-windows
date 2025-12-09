using System;
using System.Runtime.InteropServices;

namespace Benkei
{
    internal static class ImeUtility
    {
        private const int WmImeControl = 0x0283;
        private const int ImcSetopenstatus = 0x0006;
        private const int IMC_SETCONVERSIONMODE = 0x0002;
        private const int ImeCmodeNative = 0x0001;
        private const int ImeCmodeKatakana = 0x0002;
        private const int ImeCmodeFullshape = 0x0008;
        private const int ImeCmodeRoman = 0x0010;
        private const int CModeHiragana = ImeCmodeNative | ImeCmodeFullshape | ImeCmodeRoman;
        private const int CModeKatakana = ImeCmodeNative | ImeCmodeFullshape | ImeCmodeKatakana;

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

        private static bool TryGetDefaultContext(out IntPtr defaultContext)
        {
            var foreground = GetForegroundWindow();
            if (foreground == IntPtr.Zero)
            {
                defaultContext = IntPtr.Zero;
                return false;
            }

            defaultContext = ImmGetDefaultIMEWnd(foreground);
            return defaultContext != IntPtr.Zero;
        }

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("imm32.dll")]
        private static extern IntPtr ImmGetDefaultIMEWnd(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);
    }
}
