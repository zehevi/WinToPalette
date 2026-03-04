using System;
using System.Runtime.InteropServices;

namespace WinToPalette.Interception
{
    /// <summary>
    /// P/Invoke bindings for the Interception kernel driver
    /// Allows low-level keyboard and mouse input interception
    /// </summary>
    public static class InterceptionNative
    {
        private const string InterceptionDll = "interception.dll";

        [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
        public delegate int InterceptionPredicate(int device);

        // Device types
        public const int INTERCEPTION_KEYBOARD = 1;
        public const int INTERCEPTION_MOUSE = 2;

        // Keyboard scan codes
        public const ushort VK_LWIN = 0x5B;  // Left Windows key
        public const ushort VK_RWIN = 0x5C;  // Right Windows key

        // Key types
        public const ushort KEY_DOWN = 0x00;
        public const ushort KEY_UP = 0x01;
        public const ushort KEY_E0 = 0x02;
        public const ushort KEY_E1 = 0x04;

        // Keyboard filter constants
        public const ushort FILTER_KEY_NONE = 0x0000;
        public const ushort FILTER_KEY_ALL = 0xFFFF;

        // Mouse buttons
        public const ushort MOUSE_BUTTON_1_DOWN = 0x0001;
        public const ushort MOUSE_BUTTON_1_UP = 0x0002;
        public const ushort MOUSE_BUTTON_2_DOWN = 0x0004;
        public const ushort MOUSE_BUTTON_2_UP = 0x0008;

        [StructLayout(LayoutKind.Sequential)]
        public struct KeyStroke
        {
            public ushort Code;      // Scan code
            public ushort State;     // KEY_DOWN or KEY_UP
            public uint Information; // Extra information
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct MouseMotion
        {
            public short X;
            public short Y;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct MouseButtons
        {
            public ushort State;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct MouseScroll
        {
            public short Rolling;
        }

        [StructLayout(LayoutKind.Explicit, Size = 28)]
        public struct Mouse
        {
            [FieldOffset(0)]
            public MouseMotion Motion;

            [FieldOffset(0)]
            public MouseButtons Buttons;

            [FieldOffset(0)]
            public MouseScroll Scroll;

            [FieldOffset(4)]
            public uint Flags;

            [FieldOffset(8)]
            public uint Rolling;
        }

        /// <summary>
        /// Creates an interception context
        /// </summary>
        [DllImport(InterceptionDll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr interception_create_context();

        /// <summary>
        /// Destroys an interception context
        /// </summary>
        [DllImport(InterceptionDll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void interception_destroy_context(IntPtr context);

        /// <summary>
        /// Checks if device is a valid keyboard
        /// </summary>
        [DllImport(InterceptionDll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int interception_is_keyboard(int device);

        /// <summary>
        /// Checks if device is a valid mouse
        /// </summary>
        [DllImport(InterceptionDll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int interception_is_mouse(int device);

        /// <summary>
        /// Sets filter for matching devices to intercept
        /// </summary>
        [DllImport(InterceptionDll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void interception_set_filter(IntPtr context, InterceptionPredicate predicate, ushort filter);

        /// <summary>
        /// Waits for next device event
        /// </summary>
        [DllImport(InterceptionDll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int interception_wait(IntPtr context);

        /// <summary>
        /// Waits for next device event with timeout in milliseconds
        /// </summary>
        [DllImport(InterceptionDll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int interception_wait_with_timeout(IntPtr context, uint milliseconds);

        /// <summary>
        /// Retrieves next keyboard input event
        /// </summary>
        [DllImport(InterceptionDll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int interception_receive(IntPtr context, int device, ref KeyStroke stroke, uint nstroke);

        /// <summary>
        /// Waits and retrieves next input event (mouse variant)
        /// </summary>
        [DllImport(InterceptionDll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int interception_receive(IntPtr context, int device, ref Mouse mouse, uint nanos);

        /// <summary>
        /// Sends keyboard input event(s)
        /// </summary>
        [DllImport(InterceptionDll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int interception_send(IntPtr context, int device, ref KeyStroke stroke, uint nstroke);

        /// <summary>
        /// Sends an input event (mouse variant)
        /// </summary>
        [DllImport(InterceptionDll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void interception_send(IntPtr context, int device, ref Mouse mouse);

        /// <summary>
        /// Gets or sets interception filter state
        /// </summary>
        [DllImport(InterceptionDll, CallingConvention = CallingConvention.Cdecl)]
        public static extern uint interception_get_filter(IntPtr context, int device);

        /// <summary>
        /// Gets hardware identifier bytes for a device
        /// </summary>
        [DllImport(InterceptionDll, CallingConvention = CallingConvention.Cdecl)]
        public static extern uint interception_get_hardware_id(IntPtr context, int device, byte[] hardwareIdBuffer, uint bufferSize);
    }
}
