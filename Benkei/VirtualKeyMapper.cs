using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace Benkei
{
    internal static class VirtualKeyMapper
    {
        private const int VkKana = 0x15;
        private const int VkNonConvert = 0x1D;
        private const int VkOemNecEqual = 0x92;
        private const int VkOem102 = 0xE2;
        private const int VkOemYen = 0xDC;

        private static readonly Dictionary<string, int> KeyCodes = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            ["A"] = (int)Keys.A,
            ["B"] = (int)Keys.B,
            ["C"] = (int)Keys.C,
            ["D"] = (int)Keys.D,
            ["E"] = (int)Keys.E,
            ["F"] = (int)Keys.F,
            ["G"] = (int)Keys.G,
            ["H"] = (int)Keys.H,
            ["I"] = (int)Keys.I,
            ["J"] = (int)Keys.J,
            ["K"] = (int)Keys.K,
            ["L"] = (int)Keys.L,
            ["M"] = (int)Keys.M,
            ["N"] = (int)Keys.N,
            ["O"] = (int)Keys.O,
            ["P"] = (int)Keys.P,
            ["Q"] = (int)Keys.Q,
            ["R"] = (int)Keys.R,
            ["S"] = (int)Keys.S,
            ["T"] = (int)Keys.T,
            ["U"] = (int)Keys.U,
            ["V"] = (int)Keys.V,
            ["W"] = (int)Keys.W,
            ["X"] = (int)Keys.X,
            ["Y"] = (int)Keys.Y,
            ["Z"] = (int)Keys.Z,

            ["0"] = (int)Keys.D0,
            ["1"] = (int)Keys.D1,
            ["2"] = (int)Keys.D2,
            ["3"] = (int)Keys.D3,
            ["4"] = (int)Keys.D4,
            ["5"] = (int)Keys.D5,
            ["6"] = (int)Keys.D6,
            ["7"] = (int)Keys.D7,
            ["8"] = (int)Keys.D8,
            ["9"] = (int)Keys.D9,

            ["Space"] = (int)Keys.Space,
            ["Return"] = (int)Keys.Return,
            ["Enter"] = (int)Keys.Enter,
            ["Tab"] = (int)Keys.Tab,
            ["Escape"] = (int)Keys.Escape,
            ["Delete"] = (int)Keys.Delete,
            ["ForwardDelete"] = (int)Keys.Delete,
            ["Backspace"] = (int)Keys.Back,
            ["CapsLock"] = (int)Keys.CapsLock,
            ["UpArrow"] = (int)Keys.Up,
            ["DownArrow"] = (int)Keys.Down,
            ["LeftArrow"] = (int)Keys.Left,
            ["RightArrow"] = (int)Keys.Right,
            ["Home"] = (int)Keys.Home,
            ["End"] = (int)Keys.End,
            ["PageUp"] = (int)Keys.PageUp,
            ["PageDown"] = (int)Keys.PageDown,

            ["Comma"] = (int)Keys.Oemcomma,
            ["Period"] = (int)Keys.OemPeriod,
            ["Slash"] = (int)Keys.OemQuestion,
            ["Semicolon"] = (int)Keys.OemSemicolon,
            ["Quote"] = (int)Keys.OemQuotes,
            ["Minus"] = (int)Keys.OemMinus,
            ["Equal"] = (int)Keys.Oemplus,
            ["LeftBracket"] = (int)Keys.OemOpenBrackets,
            ["RightBracket"] = (int)Keys.OemCloseBrackets,
            ["Backslash"] = (int)Keys.Oem5,
            ["Grave"] = (int)Keys.Oemtilde,

            ["Shift"] = (int)Keys.ShiftKey,
            ["RightShift"] = (int)Keys.RShiftKey,
            ["Control"] = (int)Keys.ControlKey,
            ["RightControl"] = (int)Keys.RControlKey,
            ["Command"] = (int)Keys.ControlKey,
            ["RightCommand"] = (int)Keys.RControlKey,
            ["Option"] = (int)Keys.Menu,
            ["Alt"] = (int)Keys.Menu,
            ["RightOption"] = (int)Keys.RMenu,
            ["Function"] = (int)Keys.F13,

            ["IMEON"] = 0x16,
            ["IMEOFF"] = 0x1A,
            ["JIS_KeypadComma"] = VkOemNecEqual,
            ["JIS_Underscore"] = VkOem102,
            ["JIS_Yen"] = VkOemYen
        };

        private static readonly HashSet<int> NaginataKeys = new HashSet<int>(
            new[]
            {
                GetRequiredKeyCode("Q"), GetRequiredKeyCode("W"), GetRequiredKeyCode("E"), GetRequiredKeyCode("R"),
                GetRequiredKeyCode("T"), GetRequiredKeyCode("Y"), GetRequiredKeyCode("U"), GetRequiredKeyCode("I"),
                GetRequiredKeyCode("O"), GetRequiredKeyCode("P"), GetRequiredKeyCode("A"), GetRequiredKeyCode("S"),
                GetRequiredKeyCode("D"), GetRequiredKeyCode("F"), GetRequiredKeyCode("G"), GetRequiredKeyCode("H"),
                GetRequiredKeyCode("J"), GetRequiredKeyCode("K"), GetRequiredKeyCode("L"), GetRequiredKeyCode("Semicolon"),
                GetRequiredKeyCode("Z"), GetRequiredKeyCode("X"), GetRequiredKeyCode("C"), GetRequiredKeyCode("V"),
                GetRequiredKeyCode("B"), GetRequiredKeyCode("N"), GetRequiredKeyCode("M"), GetRequiredKeyCode("Comma"),
                GetRequiredKeyCode("Period"), GetRequiredKeyCode("Slash"), GetRequiredKeyCode("Space"), GetRequiredKeyCode("Return")
            }
        );

        public static int Space => GetRequiredKeyCode("Space");
        public static int Return => GetRequiredKeyCode("Return");

        public static bool TryGetKeyCode(string name, out int keyCode)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                keyCode = 0;
                return false;
            }

            return KeyCodes.TryGetValue(name.Trim(), out keyCode);
        }

        public static bool IsNaginataKey(int keyCode) => NaginataKeys.Contains(keyCode);

        public static IReadOnlyCollection<int> GetNaginataKeys() => NaginataKeys;

        private static int GetRequiredKeyCode(string name)
        {
            if (!TryGetKeyCode(name, out var keyCode))
            {
                throw new InvalidOperationException($"Virtual key mapping is missing entry for '{name}'.");
            }

            return keyCode;
        }
    }
}
