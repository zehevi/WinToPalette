using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Management;
using System.Threading;
using System.Threading.Tasks;
using WinToPalette.Logging;

namespace WinToPalette.Interception
{
    /// <summary>
    /// Manages Interception context and event handling
    /// </summary>
    public class InterceptionManager : IDisposable
    {
        private static readonly InterceptionNative.InterceptionPredicate KeyboardPredicate =
            device => InterceptionNative.interception_is_keyboard(device);

        private IntPtr _context;
        private bool _initialized;
        private CancellationTokenSource _cancellationTokenSource;
        private ILogger _logger;
        private bool _windowsKeyPending;
        private bool _windowsKeySent;
        private int _windowsKeyDevice;
        private InterceptionNative.KeyStroke _windowsKeyDownStroke;
        private string _lastKeyboardTopologySignature = string.Empty;
        private int _contextRecreateRequested;
        private string _contextRecreateReason = string.Empty;
        private const int MaxRetryAttempts = 5;
        private const int InitialRetryDelayMs = 500;

        public event EventHandler<KeyEventArgs> KeyDown;
        public event EventHandler<KeyEventArgs> KeyUp;

        public InterceptionManager(ILogger logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        /// <summary>
        /// Initializes the Interception context with automatic recovery and retry logic
        /// </summary>
        public bool Initialize()
        {
            int retryAttempt = 0;
            int retryDelayMs = InitialRetryDelayMs;

            while (retryAttempt < MaxRetryAttempts)
            {
                try
                {
                    if (retryAttempt > 0)
                    {
                        _logger.LogInfo($"Interception initialization attempt {retryAttempt + 1}/{MaxRetryAttempts}. Waiting {retryDelayMs}ms before retry...");
                        Thread.Sleep(retryDelayMs);
                    }

                    if (CreateAndConfigureContext())
                    {
                        _initialized = true;
                        _logger.LogInfo("Interception initialization complete");
                        return true;
                    }

                    // Context creation failed, prepare for retry
                    _logger.LogWarning($"Failed to create Interception context (attempt {retryAttempt + 1}/{MaxRetryAttempts})");
                    
                    if (retryAttempt < MaxRetryAttempts - 1)
                    {
                        // Try to recover the driver before next attempt
                        if (retryAttempt == 0)
                        {
                            _logger.LogInfo("Attempting driver recovery...");
                            AttemptDriverRecovery();
                        }
                        
                        // Exponential backoff: 500ms, 1s, 2s, 4s
                        retryDelayMs = Math.Min(retryDelayMs * 2, 4000);
                    }

                    retryAttempt++;
                }
                catch (DllNotFoundException ex)
                {
                    _logger.LogError($"Interception.dll not found: {ex.Message}");
                    return false;
                }
                catch (Exception ex)
                {
                    _logger.LogError($"Exception during initialization: {ex.Message}");
                    retryAttempt++;
                }
            }

            _logger.LogError($"Failed to initialize Interception after {MaxRetryAttempts} attempts. Is the Interception driver installed and loaded?");
            return false;
        }

        /// <summary>
        /// Starts the event listening loop with resilience to device disconnections
        /// </summary>
        public async Task StartListeningAsync()
        {
            if (!_initialized || _context == IntPtr.Zero)
            {
                _logger.LogError("Interception manager not initialized");
                return;
            }

            _cancellationTokenSource = new CancellationTokenSource();

            await Task.Run(() =>
            {
                _logger.LogInfo("Starting Interception event listening");

                // Log all detected keyboard devices at startup
                LogDetectedKeyboards();
                _lastKeyboardTopologySignature = GetKeyboardTopologySignature();

                DateTime lastTopologyCheckTime = DateTime.Now;
                const int topologyCheckIntervalMs = 2000;
                int errorCount = 0;
                const int maxConsecutiveErrors = 10;

                while (!_cancellationTokenSource.Token.IsCancellationRequested)
                {
                    try
                    {
                        if (Interlocked.Exchange(ref _contextRecreateRequested, 0) == 1)
                        {
                            string reason = string.IsNullOrWhiteSpace(_contextRecreateReason)
                                ? "External context recreate request"
                                : _contextRecreateReason;

                            TryRecreateContext(reason);
                            _lastKeyboardTopologySignature = GetKeyboardTopologySignature();
                            lastTopologyCheckTime = DateTime.Now;
                        }

                        if (_context == IntPtr.Zero)
                        {
                            if (!TryRecreateContext("Context is null during listen loop"))
                            {
                                Thread.Sleep(250);
                                continue;
                            }

                            _lastKeyboardTopologySignature = GetKeyboardTopologySignature();
                            lastTopologyCheckTime = DateTime.Now;
                        }

                        if ((DateTime.Now - lastTopologyCheckTime).TotalMilliseconds >= topologyCheckIntervalMs)
                        {
                            string currentSignature = GetKeyboardTopologySignature();
                            if (!string.Equals(currentSignature, _lastKeyboardTopologySignature, StringComparison.Ordinal))
                            {
                                _logger.LogWarning("Keyboard device topology changed. Recreating interception context...");
                                if (TryRecreateContext("Keyboard topology changed after USB reconnect"))
                                {
                                    _lastKeyboardTopologySignature = GetKeyboardTopologySignature();
                                }
                            }

                            lastTopologyCheckTime = DateTime.Now;
                        }

                        // Wait for any keyboard input (handles dynamic device discovery)
                        int device = InterceptionNative.interception_wait_with_timeout(_context, 100);

                        // No input on any device
                        if (device <= 0)
                        {
                            errorCount = 0;
                            continue;
                        }

                        // Verify it's a keyboard - log non-keyboards too
                        int isKeyboard = InterceptionNative.interception_is_keyboard(device);
                        if (isKeyboard == 0)
                        {
                            _logger.LogDebug($"Device {device} returned from wait, but is not a keyboard");
                            continue;
                        }

                        var stroke = new InterceptionNative.KeyStroke();
                        int result = InterceptionNative.interception_receive(_context, device, ref stroke, 1);

                        if (result > 0)
                        {
                            errorCount = 0;
                            _logger.LogDebug($"Key event from device {device}: Code=0x{stroke.Code:X4}, State={stroke.State}");
                            HandleKeyEvent(device, stroke);
                        }
                        else if (result == 0)
                        {
                            // No data available (shouldn't happen with wait_with_timeout, but log it)
                            _logger.LogDebug($"Device {device} signaled but no data available (result=0)");
                        }
                        else if (result < 0)
                        {
                            // Device disconnected or error occurred
                            errorCount++;
                            _logger.LogDebug($"Interception receive error on device {device}: {result}");

                            // If too many errors, break to prevent infinite error loop
                            if (errorCount >= maxConsecutiveErrors)
                            {
                                _logger.LogError($"Too many consecutive interception errors. Recreating context...");
                                TryRecreateContext("Too many consecutive receive errors");
                                _lastKeyboardTopologySignature = GetKeyboardTopologySignature();
                                errorCount = 0;
                                lastTopologyCheckTime = DateTime.Now;
                            }

                            // Brief sleep before retrying
                            Thread.Sleep(50);
                        }
                    }
                    catch (OperationCanceledException)
                    {
                        _logger.LogInfo("Interception listening cancelled");
                        break;
                    }
                    catch (Exception ex)
                    {
                        errorCount++;
                        _logger.LogError($"Error in interception loop: {ex.Message}");

                        if (errorCount >= maxConsecutiveErrors)
                        {
                            _logger.LogError($"Too many errors in interception loop. Recreating context...");
                            TryRecreateContext("Too many loop exceptions");
                            _lastKeyboardTopologySignature = GetKeyboardTopologySignature();
                            errorCount = 0;
                            lastTopologyCheckTime = DateTime.Now;
                        }

                        Thread.Sleep(100);
                    }
                }

                _logger.LogInfo("Interception event listening stopped");
            });
        }

        private bool CreateAndConfigureContext()
        {
            try
            {
                _context = InterceptionNative.interception_create_context();
                if (_context == IntPtr.Zero)
                {
                    _logger.LogDebug("interception_create_context returned null pointer");
                    return false;
                }

                _logger.LogInfo("Interception context created successfully");

                InterceptionNative.interception_set_filter(
                    _context,
                    KeyboardPredicate,
                    InterceptionNative.FILTER_KEY_ALL);

                _logger.LogInfo("Keyboard filter configured (all key events)");
                _logger.LogInfo("Ready to intercept keyboard events");
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError($"Failed to create/configure interception context: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Attempts to recover the Interception driver by restarting related services
        /// </summary>
        private void AttemptDriverRecovery()
        {
            try
            {
                _logger.LogInfo("Attempting to recover Interception driver...");
                
                // Try to restart keyboard and mouse driver services
                string[] serviceNames = { "keyboard", "mouse" };
                
                foreach (var serviceName in serviceNames)
                {
                    try
                    {
                        var service = new ManagementObject($"Win32_SystemDriver.Name='{serviceName}'");
                        service.InvokeMethod("StartService", null);
                        _logger.LogInfo($"Attempted to start {serviceName} service");
                    }
                    catch (Exception ex)
                    {
                        _logger.LogDebug($"Could not start {serviceName} service: {ex.Message}");
                    }
                }

                // Give services time to start
                Thread.Sleep(1000);
                _logger.LogInfo("Driver recovery attempt completed");
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Driver recovery failed: {ex.Message}");
            }
        }

        private bool TryRecreateContext(string reason)
        {
            try
            {
                _logger.LogInfo($"Recreating interception context: {reason}");

                if (_context != IntPtr.Zero)
                {
                    InterceptionNative.interception_destroy_context(_context);
                    _context = IntPtr.Zero;
                }

                ResetWindowsKeyState();

                // Ask Windows to re-enumerate all keyboard devices before binding the new context.
                // This releases any stuck USB keyboard state left over from interception filter churn.
                TriggerPnpKeyboardRescan();

                if (!CreateAndConfigureContext())
                {
                    _logger.LogError("Context recreation failed");
                    return false;
                }

                LogDetectedKeyboards();
                _logger.LogInfo("Interception context recreation complete");
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError($"Failed to recreate context: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Triggers a Windows PnP device re-scan for keyboard class devices.
        /// This forces re-enumeration of USB keyboards that may have been left in a stuck
        /// state by interception filter churn during process restarts or context recreation.
        /// </summary>
        private void TriggerPnpKeyboardRescan()
        {
            try
            {
                _logger.LogInfo("Triggering PnP keyboard device rescan...");

                // pnputil /scan-devices is the safest way to force Windows to re-enumerate
                // all devices without needing to know specific instance IDs.
                var psi = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = "pnputil.exe",
                    Arguments = "/scan-devices",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };

                using var proc = System.Diagnostics.Process.Start(psi);
                proc?.WaitForExit(3000);
                _logger.LogInfo("PnP keyboard device rescan complete");
            }
            catch (Exception ex)
            {
                _logger.LogDebug($"PnP rescan skipped: {ex.Message}");
            }
        }

        private void LogDetectedKeyboards()
        {
            try
            {
                _logger.LogInfo("Scanning for keyboard devices...");
                int keyboardSlotCount = 0;
                int activeKeyboardCount = 0;
                byte[] buffer = new byte[512];
                
                for (int i = 1; i <= 20; i++)
                {
                    if (InterceptionNative.interception_is_keyboard(i) != 0)
                    {
                        keyboardSlotCount++;

                        Array.Clear(buffer, 0, buffer.Length);
                        uint length = InterceptionNative.interception_get_hardware_id(_context, i, buffer, (uint)buffer.Length);
                        if (length > 0)
                        {
                            activeKeyboardCount++;
                            _logger.LogInfo($"Active keyboard detected at index {i}");
                        }
                    }
                }

                _logger.LogInfo($"Keyboard slots detected: {keyboardSlotCount}, active keyboards with hardware IDs: {activeKeyboardCount}");
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error scanning for keyboards: {ex.Message}");
            }
        }

        private string GetKeyboardTopologySignature()
        {
            if (_context == IntPtr.Zero)
            {
                return string.Empty;
            }

            var signatures = new List<string>();
            byte[] buffer = new byte[512];

            for (int device = 1; device <= 20; device++)
            {
                if (InterceptionNative.interception_is_keyboard(device) == 0)
                {
                    continue;
                }

                Array.Clear(buffer, 0, buffer.Length);
                uint length = InterceptionNative.interception_get_hardware_id(_context, device, buffer, (uint)buffer.Length);

                if (length > 0)
                {
                    int safeLength = (int)Math.Min(length, (uint)buffer.Length);
                    string hardwareBytes = Convert.ToBase64String(buffer, 0, safeLength);
                    signatures.Add(hardwareBytes);
                }
            }

            signatures.Sort(StringComparer.Ordinal);
            return string.Join("|", signatures);
        }

        /// <summary>
        /// Stops the event listening loop
        /// </summary>
        public void StopListening()
        {
            _cancellationTokenSource?.Cancel();
        }

        private void HandleKeyEvent(int device, InterceptionNative.KeyStroke stroke)
        {
            bool isKeyUp = (stroke.State & InterceptionNative.KEY_UP) == InterceptionNative.KEY_UP;
            bool isKeyDown = !isKeyUp;
            bool isWindowsKey = stroke.Code == InterceptionNative.VK_LWIN;

            if (isWindowsKey)
            {
                HandleWindowsKeyEvent(device, stroke, isKeyDown, isKeyUp);
                return;
            }

            if (_windowsKeyPending && !_windowsKeySent)
            {
                InterceptionNative.interception_send(_context, _windowsKeyDevice, ref _windowsKeyDownStroke, 1);
                _windowsKeySent = true;
            }

            // For all other keys, pass them through to the OS
            InterceptionNative.interception_send(_context, device, ref stroke, 1);
        }

        private void HandleWindowsKeyEvent(int device, InterceptionNative.KeyStroke stroke, bool isKeyDown, bool isKeyUp)
        {
            if (isKeyDown)
            {
                _windowsKeyPending = true;
                _windowsKeySent = false;
                _windowsKeyDevice = device;
                _windowsKeyDownStroke = stroke;

                _logger.LogDebug($"Windows key down queued: {(stroke.Code == InterceptionNative.VK_LWIN ? "Left" : "Right")}");
                return;
            }

            if (isKeyUp)
            {
                bool sameWindowsKey = _windowsKeyPending && stroke.Code == _windowsKeyDownStroke.Code;

                if (sameWindowsKey && !_windowsKeySent)
                {
                    _logger.LogDebug($"Standalone Windows key released: {(stroke.Code == InterceptionNative.VK_LWIN ? "Left" : "Right")}");
                    var keyUpTimestamp = Stopwatch.GetTimestamp();
                    KeyUp?.Invoke(this, new KeyEventArgs { Key = stroke.Code, IsKeyDown = false, Timestamp = keyUpTimestamp });
                    ResetWindowsKeyState();
                    return;
                }

                if (sameWindowsKey && _windowsKeySent)
                {
                    InterceptionNative.interception_send(_context, device, ref stroke, 1);
                    ResetWindowsKeyState();
                    return;
                }
            }

            InterceptionNative.interception_send(_context, device, ref stroke, 1);
        }

        private void ResetWindowsKeyState()
        {
            _windowsKeyPending = false;
            _windowsKeySent = false;
            _windowsKeyDevice = 0;
            _windowsKeyDownStroke = default;
        }

        public void RequestContextRecreate(string reason)
        {
            _contextRecreateReason = reason ?? "External request";
            Interlocked.Exchange(ref _contextRecreateRequested, 1);
        }

        public void Dispose()
        {
            StopListening();
            _cancellationTokenSource?.Dispose();

            if (_context != IntPtr.Zero)
            {
                InterceptionNative.interception_destroy_context(_context);
                _context = IntPtr.Zero;
            }

            ResetWindowsKeyState();

            _initialized = false;
        }
    }

    public class KeyEventArgs : EventArgs
    {
        public ushort Key { get; set; }
        public bool IsKeyDown { get; set; }
        public long Timestamp { get; set; }
    }
}
