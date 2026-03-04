using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using WinToPalette.Interception;
using WinToPalette.Logging;

namespace WinToPalette.Input
{
    public sealed class KeyboardFallbackMonitor : IDisposable
    {
        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_KEYUP = 0x0101;
        private const int WM_SYSKEYDOWN = 0x0104;
        private const int WM_SYSKEYUP = 0x0105;
        private const uint WM_QUIT = 0x0012;
        private const int VK_LWIN = 0x5B;
        private const uint KEYEVENTF_EXTENDEDKEY = 0x0001;
        private const uint KEYEVENTF_KEYUP = 0x0002;
        private static readonly IntPtr INJECTED_EVENT_MARKER = new IntPtr(0x5A4D5243); // Magic number to identify our injections

        private readonly ILogger _logger;
        private readonly AutoResetEvent _startedEvent = new AutoResetEvent(false);
        private Thread _hookThread;
        private uint _hookThreadId;
        private IntPtr _hookHandle;
        private LowLevelKeyboardProc _proc;
        private volatile bool _running;
        private bool _leftWinDown;
        private bool _comboDetected;
        private int _comboKey; // Track which key was part of combo to suppress its up event

        public event EventHandler<KeyEventArgs> LeftWindowsKeyReleased;

        public KeyboardFallbackMonitor(ILogger logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        public void Start()
        {
            if (_running)
            {
                return;
            }

            _running = true;
            _hookThread = new Thread(HookThreadMain)
            {
                IsBackground = true,
                Name = "KeyboardFallbackMonitorThread"
            };
            _hookThread.Start();

            if (!_startedEvent.WaitOne(TimeSpan.FromSeconds(5)))
            {
                _logger.LogWarning("Keyboard fallback monitor did not start within timeout");
            }
        }

        public void Stop()
        {
            if (!_running)
            {
                return;
            }

            _running = false;

            if (_hookThreadId != 0)
            {
                PostThreadMessage(_hookThreadId, WM_QUIT, IntPtr.Zero, IntPtr.Zero);
            }

            if (_hookThread != null && _hookThread.IsAlive)
            {
                _hookThread.Join(1000);
            }
        }

        private void HookThreadMain()
        {
            try
            {
                _hookThreadId = GetCurrentThreadId();
                _proc = HookCallback;

                IntPtr moduleHandle = GetModuleHandle(Process.GetCurrentProcess().MainModule?.ModuleName);
                _hookHandle = SetWindowsHookEx(WH_KEYBOARD_LL, _proc, moduleHandle, 0);

                if (_hookHandle == IntPtr.Zero)
                {
                    _logger.LogWarning("Failed to install keyboard fallback hook");
                    _startedEvent.Set();
                    return;
                }

                _logger.LogInfo("Keyboard fallback monitor started");
                _startedEvent.Set();

                while (_running && GetMessage(out MSG msg, IntPtr.Zero, 0, 0) > 0)
                {
                    TranslateMessage(ref msg);
                    DispatchMessage(ref msg);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"Keyboard fallback monitor error: {ex.Message}");
            }
            finally
            {
                if (_hookHandle != IntPtr.Zero)
                {
                    UnhookWindowsHookEx(_hookHandle);
                    _hookHandle = IntPtr.Zero;
                }

                _hookThreadId = 0;
                _logger.LogInfo("Keyboard fallback monitor stopped");
            }
        }

        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                int message = wParam.ToInt32();
                KBDLLHOOKSTRUCT kbData = Marshal.PtrToStructure<KBDLLHOOKSTRUCT>(lParam);
                int vkCode = (int)kbData.vkCode;

                // Let our injected events pass through
                if (kbData.dwExtraInfo == INJECTED_EVENT_MARKER)
                {
                    return CallNextHookEx(_hookHandle, nCode, wParam, lParam);
                }

                bool isKeyDown = message == WM_KEYDOWN || message == WM_SYSKEYDOWN;
                bool isKeyUp = message == WM_KEYUP || message == WM_SYSKEYUP;

                if (vkCode == VK_LWIN)
                {
                    if (isKeyDown)
                    {
                        _leftWinDown = true;
                        _comboDetected = false;
                        _comboKey = 0;
                        // Pass through Win down to let Windows see it
                        return CallNextHookEx(_hookHandle, nCode, wParam, lParam);
                    }
                    else if (isKeyUp)
                    {
                        if (_leftWinDown && !_comboDetected)
                        {
                            // Standalone Win key - fire our event
                            var timestamp = Stopwatch.GetTimestamp();
                            LeftWindowsKeyReleased?.Invoke(this, new KeyEventArgs
                            {
                                Key = (ushort)VK_LWIN,
                                IsKeyDown = false,
                                Timestamp = timestamp
                            });
                            
                            // Suppress Win up to prevent Start menu for standalone press
                            _leftWinDown = false;
                            _comboDetected = false;
                            _comboKey = 0;
                            return (IntPtr)1;
                        }
                        else
                        {
                            // Combo detected - pass through Win up so combo completes normally
                            _leftWinDown = false;
                            _comboDetected = false;
                            _comboKey = 0;
                            return CallNextHookEx(_hookHandle, nCode, wParam, lParam);
                        }
                    }
                }
                else if (_leftWinDown && isKeyDown && !_comboDetected)
                {
                    // Combo detected - pass through the other key
                    _comboDetected = true;
                    _comboKey = vkCode;
                    return CallNextHookEx(_hookHandle, nCode, wParam, lParam);
                }
            }

            return CallNextHookEx(_hookHandle, nCode, wParam, lParam);
        }

        private void InjectWinKeyCombo(ushort vkCode)
        {
            try
            {
                _logger.LogInfo($"Injecting Win+{vkCode:X} combo");
                
                // Send full key sequence: Win down, Key down, Key up, Win up
                INPUT[] inputs = new INPUT[4];
                
                // Win key down
                inputs[0] = new INPUT
                {
                    type = 1, // INPUT_KEYBOARD
                    ki = new KEYBDINPUT
                    {
                        wVk = VK_LWIN,
                        wScan = 0,
                        dwFlags = KEYEVENTF_EXTENDEDKEY,
                        time = 0,
                        dwExtraInfo = INJECTED_EVENT_MARKER
                    }
                };
                
                // Other key down
                inputs[1] = new INPUT
                {
                    type = 1, // INPUT_KEYBOARD
                    ki = new KEYBDINPUT
                    {
                        wVk = vkCode,
                        wScan = 0,
                        dwFlags = 0,
                        time = 0,
                        dwExtraInfo = INJECTED_EVENT_MARKER
                    }
                };

                // Other key up
                inputs[2] = new INPUT
                {
                    type = 1, // INPUT_KEYBOARD
                    ki = new KEYBDINPUT
                    {
                        wVk = vkCode,
                        wScan = 0,
                        dwFlags = KEYEVENTF_KEYUP,
                        time = 0,
                        dwExtraInfo = INJECTED_EVENT_MARKER
                    }
                };

                // Win key up
                inputs[3] = new INPUT
                {
                    type = 1, // INPUT_KEYBOARD
                    ki = new KEYBDINPUT
                    {
                        wVk = VK_LWIN,
                        wScan = 0,
                        dwFlags = KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP,
                        time = 0,
                        dwExtraInfo = INJECTED_EVENT_MARKER
                    }
                };

                uint result = SendInput(4, inputs, Marshal.SizeOf(typeof(INPUT)));
                if (result != 4)
                {
                    _logger.LogWarning($"SendInput returned {result}, expected 4");
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Failed to inject Win key combo: {ex.Message}");
            }
        }

        public void Dispose()
        {
            Stop();
            _startedEvent.Dispose();
        }

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        private struct KBDLLHOOKSTRUCT
        {
            public uint vkCode;
            public uint scanCode;
            public uint flags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Explicit)]
        private struct INPUT
        {
            [FieldOffset(0)]
            public uint type;
            [FieldOffset(4)]
            public MOUSEINPUT mi;
            [FieldOffset(4)]
            public KEYBDINPUT ki;
            [FieldOffset(4)]
            public HARDWAREINPUT hi;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MOUSEINPUT
        {
            public int dx;
            public int dy;
            public uint mouseData;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct KEYBDINPUT
        {
            public ushort wVk;
            public ushort wScan;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct HARDWAREINPUT
        {
            public uint uMsg;
            public ushort wParamL;
            public ushort wParamH;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MSG
        {
            public IntPtr hwnd;
            public uint message;
            public IntPtr wParam;
            public IntPtr lParam;
            public uint time;
            public POINT pt;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int x;
            public int y;
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("user32.dll")]
        private static extern sbyte GetMessage(out MSG lpMsg, IntPtr hWnd, uint wMsgFilterMin, uint wMsgFilterMax);

        [DllImport("user32.dll")]
        private static extern bool TranslateMessage(ref MSG lpMsg);

        [DllImport("user32.dll")]
        private static extern IntPtr DispatchMessage(ref MSG lpMsg);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool PostThreadMessage(uint idThread, uint msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();
    }
}
