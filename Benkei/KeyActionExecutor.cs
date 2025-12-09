using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Benkei
{
    internal sealed class KeyActionExecutor
    {
        private readonly Action<bool> _setRepeatAllowed;
        private readonly Action _resetRequest;
        private readonly HashSet<ushort> _latchedKeys = new HashSet<ushort>();
        private static readonly IntPtr BenkeiMarker = new IntPtr(0x42454E4B); // "BENK"

        public KeyActionExecutor(Action<bool> setRepeatAllowed, Action resetRequest)
        {
            _setRepeatAllowed = setRepeatAllowed ?? throw new ArgumentNullException(nameof(setRepeatAllowed));
            _resetRequest = resetRequest ?? throw new ArgumentNullException(nameof(resetRequest));
        }

        public void Execute(IEnumerable<NaginataAction> actions)
        {
            if (actions == null)
            {
                return;
            }

            foreach (var action in actions)
            {
                switch (action.Kind)
                {
                    case NaginataActionType.Tap:
                        if (TryGetKey(action.Value, out var tapKey))
                        {
                            TapKey(tapKey);
                        }
                        else
                        {
                            Console.WriteLine($"[Benkei] Unknown tap key '{action.Value}'.");
                        }

                        break;
                    case NaginataActionType.Press:
                        if (TryGetKey(action.Value, out var pressKey))
                        {
                            PressKey(pressKey);
                        }
                        else
                        {
                            Console.WriteLine($"[Benkei] Unknown press key '{action.Value}'.");
                        }

                        break;
                    case NaginataActionType.Release:
                        if (TryGetKey(action.Value, out var releaseKey))
                        {
                            ReleaseKey(releaseKey);
                        }
                        else
                        {
                            Console.WriteLine($"[Benkei] Unknown release key '{action.Value}'.");
                        }

                        break;
                    case NaginataActionType.Character:
                        SendCharacters(action.Value);
                        break;
                    case NaginataActionType.Repeat:
                        var allowRepeat = bool.TryParse(action.Value, out var parsed) && parsed;
                        _setRepeatAllowed(allowRepeat);
                        break;
                    case NaginataActionType.Reset:
                        ReleaseLatchedKeys();
                        _setRepeatAllowed(false);
                        _resetRequest();
                        break;
                }
            }
        }

        public void ReleaseLatchedKeys()
        {
            foreach (var key in _latchedKeys.ToArray())
            {
                SendKey(key, true);
                _latchedKeys.Remove(key);
            }
        }

        private static bool TryGetKey(string name, out ushort keyCode)
        {
            keyCode = 0;
            if (!VirtualKeyMapper.TryGetKeyCode(name, out var code))
            {
                return false;
            }

            keyCode = (ushort)code;
            return true;
        }

        public void TapKey(ushort key)
        {
            SendKey(key, false);
            SendKey(key, true);
        }

        public void PressKey(ushort key)
        {
            if (_latchedKeys.Add(key))
            {
                SendKey(key, false);
            }
        }

        public void ReleaseKey(ushort key)
        {
            _latchedKeys.Remove(key);
            SendKey(key, true);
        }

        private void SendCharacters(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return;
            }

            // 薙刀式リセット
            // 未変換文字の確定
            if (ImeUtility.TryHasUnconvertedText())
            {
                TapKey((ushort)Keys.Return);
            }
            ImeUtility.TryTurnOff();
            foreach (var ch in value)
            {
                SendUnicode((ushort)ch, false);
                SendUnicode((ushort)ch, true);
                System.Threading.Thread.Sleep(10);
            }
            ImeUtility.TryTurnOnHiragana();
        }

        private static void SendKey(ushort keyCode, bool keyUp)
        {
            var input = new INPUT
            {
                type = InputKeyboard,
                U = new InputUnion
                {
                    ki = new KEYBDINPUT
                    {
                        wVk = keyCode,
                        wScan = 0,
                        dwFlags = keyUp ? KeyeventfKeyup : 0,
                        dwExtraInfo = BenkeiMarker
                    }
                }
            };

            var cbSize = Marshal.SizeOf(typeof(INPUT));
            Console.WriteLine($"[SendKey] Sending key: {keyCode}, keyUp: {keyUp}, cbSize: {cbSize}");
            var result = SendInput(1, new[] { input }, cbSize);
            if (result == 0)
            {
                var error = Marshal.GetLastWin32Error();
                Debug.WriteLine($"[SendKey] FAILED! Error code: {error}");
                Console.WriteLine($"[Benkei] SendInput failed: key={keyCode}, error={error}");
            }
            else
            {
                Debug.WriteLine($"[SendKey] SUCCESS! Sent {result} event(s)");
            }
        }

        private static void SendUnicode(ushort rune, bool keyUp)
        {
            var input = new INPUT
            {
                type = InputKeyboard,
                U = new InputUnion
                {
                    ki = new KEYBDINPUT
                    {
                        wVk = 0,
                        wScan = rune,
                        dwFlags = keyUp ? KeyeventfUnicode | KeyeventfKeyup : KeyeventfUnicode,
                        dwExtraInfo = BenkeiMarker
                    }
                }
            };

            var cbSize = Marshal.SizeOf(typeof(INPUT));
            Debug.WriteLine($"[SendUnicode] Sending char: '{(char)rune}' (U+{rune:X4}), keyUp: {keyUp}, cbSize: {cbSize}");
            var result = SendInput(1, new[] { input }, cbSize);
            if (result == 0)
            {
                var error = Marshal.GetLastWin32Error();
                Debug.WriteLine($"[SendUnicode] FAILED! Error code: {error}");
                Console.WriteLine($"[Benkei] SendInput failed: char='{(char)rune}', error={error}");
            }
            else
            {
                Debug.WriteLine($"[SendUnicode] SUCCESS! Sent {result} event(s)");
            }
        }

        private const uint InputKeyboard = 1;
        private const uint KeyeventfKeyup = 0x0002;
        private const uint KeyeventfUnicode = 0x0004;

        [StructLayout(LayoutKind.Sequential)]
        private struct INPUT
        {
            public uint type;
            public InputUnion U;
        }

        // [StructLayout(LayoutKind.Explicit)]
        // private struct InputUnion
        // {
        //     [FieldOffset(0)] public KEYBDINPUT ki;
        // }

        // [StructLayout(LayoutKind.Sequential)]
        // private struct KEYBDINPUT
        // {
        //     public ushort wVk;
        //     public ushort wScan;
        //     public uint dwFlags;
        //     public uint time;
        //     public IntPtr dwExtraInfo;
        // }

        [StructLayout(LayoutKind.Sequential)]
        private struct MOUSEINPUT {
            public int dx;
            public int dy;
            public int mouseData;
            public int dwFlags;
            public int time;
            public IntPtr dwExtraInfo;
        };

        [StructLayout(LayoutKind.Sequential)]
        private struct KEYBDINPUT {
            public ushort wVk;
            public ushort wScan;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        };

        [StructLayout(LayoutKind.Sequential)]
        private struct HARDWAREINPUT {
            public int uMsg;
            public short wParamL;
            public short wParamH;
        };

        [StructLayout(LayoutKind.Explicit)]
        private struct InputUnion {
            [FieldOffset(0)]
            public int type;
            [FieldOffset(0)]
            public MOUSEINPUT no;
            [FieldOffset(0)]
            public KEYBDINPUT ki;
            [FieldOffset(0)]
            public HARDWAREINPUT hi;
        };

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

        [DllImport("user32.dll")]
        private static extern IntPtr GetMessageExtraInfo();
    }
}
