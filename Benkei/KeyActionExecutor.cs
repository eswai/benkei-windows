using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace Benkei
{
    internal sealed class KeyActionExecutor
    {
        private readonly Action<bool> _setRepeatAllowed;
        private readonly Action _resetRequest;
        private readonly HashSet<ushort> _latchedKeys = new HashSet<ushort>();
        private static readonly IntPtr BenkeiMarker = new IntPtr(0x42454E4B); // "BENK"
        private static readonly int InputStructSize = Marshal.SizeOf(typeof(INPUT));

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

            var tapBuffer = new List<INPUT>();
            foreach (var action in actions)
            {
                switch (action.Kind)
                {
                    case NaginataActionType.Tap:
                        if (TryGetKey(action.Value, out var tapKey))
                        {
                            tapBuffer.Add(CreateKeyboardInput(tapKey, false));
                            tapBuffer.Add(CreateKeyboardInput(tapKey, true));
                        }
                        else
                        {
                            Logger.Log($"[Benkei] Unknown tap key '{action.Value}'.");
                        }

                        break;
                    case NaginataActionType.Press:
                        if (TryGetKey(action.Value, out var pressKey))
                        {
                            tapBuffer.Add(CreateKeyboardInput(pressKey, false));
                        }
                        else
                        {
                            Logger.Log($"[Benkei] Unknown press key '{action.Value}'.");
                        }

                        break;
                    case NaginataActionType.Release:
                        if (TryGetKey(action.Value, out var releaseKey))
                        {
                            tapBuffer.Add(CreateKeyboardInput(releaseKey, true));
                        }
                        else
                        {
                            Logger.Log($"[Benkei] Unknown release key '{action.Value}'.");
                        }

                        break;
                    case NaginataActionType.Character:
                        FlushTapBuffer(tapBuffer);
                        SendCharacters(action.Value);
                        break;
                    case NaginataActionType.Repeat:
                        FlushTapBuffer(tapBuffer);
                        var allowRepeat = bool.TryParse(action.Value, out var parsed) && parsed;
                        _setRepeatAllowed(allowRepeat);
                        break;
                    case NaginataActionType.Reset:
                        FlushTapBuffer(tapBuffer);
                        ReleaseLatchedKeys();
                        _setRepeatAllowed(false);
                        _resetRequest();
                        break;
                }
            }

            FlushTapBuffer(tapBuffer);
        }

        public void ReleaseLatchedKeys()
        {
            if (_latchedKeys.Count == 0)
            {
                return;
            }

            var releases = new INPUT[_latchedKeys.Count];
            var index = 0;
            foreach (var key in _latchedKeys)
            {
                releases[index++] = CreateKeyboardInput(key, true);
            }

            _latchedKeys.Clear();
            SendInputs(releases, "release latched keys");
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
            var tapBuffer = new List<INPUT>();
            tapBuffer.Add(CreateKeyboardInput(key, false));
            tapBuffer.Add(CreateKeyboardInput(key, true));
            FlushTapBuffer(tapBuffer);
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
            var unicodeInputs = new List<INPUT>(value.Length * 2);
            foreach (var ch in value)
            {
                unicodeInputs.Add(CreateUnicodeInput((ushort)ch, false));
                unicodeInputs.Add(CreateUnicodeInput((ushort)ch, true));
            }

            if (unicodeInputs.Count > 0)
            {
                SendInputs(unicodeInputs.ToArray(), $"unicode batch count={value.Length}");
            }
        }

        private void FlushTapBuffer(List<INPUT> tapBuffer)
        {
            if (tapBuffer == null || tapBuffer.Count == 0)
            {
                return;
            }

            SendInputs(tapBuffer.ToArray(), $"tap buffer count={tapBuffer.Count}");
            tapBuffer.Clear();
        }

        private static INPUT CreateKeyboardInput(ushort keyCode, bool keyUp)
        {
            return new INPUT
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
        }

        private static INPUT CreateUnicodeInput(ushort rune, bool keyUp)
        {
            return new INPUT
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
        }

        private static void SendInputs(INPUT[] inputs, string context)
        {
            if (inputs == null || inputs.Length == 0)
            {
                return;
            }

            Logger.Log($"[SendKey] Sending {inputs.Length} event(s): {context}, cbSize: {InputStructSize}");
            var result = SendInput((uint)inputs.Length, inputs, InputStructSize);
            if (result == 0)
            {
                var error = Marshal.GetLastWin32Error();
                Debug.WriteLine($"[SendKey] FAILED ({context}) Error code: {error}");
                Logger.Log($"[Benkei] SendInput failed ({context}): error={error}");
            }
            else
            {
                Debug.WriteLine($"[SendKey] SUCCESS ({context}) Sent {result} event(s)");
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
