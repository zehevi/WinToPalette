using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Management;
using System.Threading;
using System.Threading.Tasks;
using WinToPalette.Input;
using WinToPalette.Interception;
using WinToPalette.Logging;
using WinToPalette.Notifications;
using WinToPalette.PowerToys;
using WinToPalette.Startup;

namespace WinToPalette
{
    public class Program
    {
        private static ILogger _logger;
        private static InterceptionManager _interceptionManager;
        private static KeyboardFallbackMonitor _keyboardFallbackMonitor;
        private static PowerToysLauncher _powerToysLauncher;
        private static PowerToysNotificationService _powerToysNotificationService;
        private static StartupManager _startupManager;
        private static ManagementEventWatcher _deviceChangeWatcher;
        private static int _deviceChangeVersion;
        private static DateTime _lastSelfRestartUtc = DateTime.MinValue;
        private static DateTime _lastPaletteLaunchUtc = DateTime.MinValue;
        private static readonly object _launchSync = new object();
        private static readonly List<(string label, long ticks)> _timingBuffer = new List<(string, long)>();

        [STAThread]
        static async Task Main(string[] args)
        {
            try
            {
                // Initialize logging
                _logger = new CompositeLogger(
                    new ConsoleLogger(),
                    new FileLogger()
                );

                bool isInteractive = Environment.UserInteractive && !Console.IsInputRedirected;

                _logger.LogInfo("=== WinToPalette Launcher Started ===");
                _logger.LogInfo($"Running in interactive mode: {isInteractive}");

                // Check for admin privileges
                if (!StartupManager.IsElevated())
                {
                    _logger.LogInfo("Application requires administrator privileges. Requesting elevation...");
                    
                    if (StartupManager.RequestElevation(args))
                    {
                        _logger.LogInfo("Elevation request sent. Current process will exit.");
                        return;
                    }
                    else
                    {
                        _logger.LogError("Failed to elevate privileges. Application requires admin rights.");
                        if (isInteractive)
                        {
                            Console.WriteLine("\nPress any key to exit...");
                            WaitForKeyPress();
                        }
                        return;
                    }
                }

                _logger.LogInfo("Running with administrator privileges");

                // Ensure Interception DLL is available
                EnsureInterceptionDll(_logger);

                // Initialize Interception
                _interceptionManager = new InterceptionManager(_logger);
                if (!_interceptionManager.Initialize())
                {
                    _logger.LogError("Failed to initialize Interception. Make sure the driver is installed.");
                    if (isInteractive)
                    {
                        Console.WriteLine("\nPress any key to exit...");
                        WaitForKeyPress();
                    }
                    return;
                }

                // Initialize PowerToys launcher
                _powerToysLauncher = new PowerToysLauncher(_logger);
                _powerToysNotificationService = new PowerToysNotificationService(_logger);
                if (!_powerToysLauncher.DetectPowerToys())
                {
                    _logger.LogWarning("PowerToys not detected. The application will wait for it to be installed.");
                }

                // Initialize startup manager
                _startupManager = new StartupManager(_logger);

                // Register for startup if requested
                if (args.Length > 0 && args[0].Equals("--register-startup", StringComparison.OrdinalIgnoreCase))
                {
                    var executablePath = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName;
                    _startupManager.RegisterForStartup(executablePath);
                    _logger.LogInfo("Registered for startup. Application will auto-launch on next boot.");
                    return;
                }

                // Hook into key events
                _interceptionManager.KeyUp += OnWindowsKeyUp;

                _keyboardFallbackMonitor = new KeyboardFallbackMonitor(_logger);
                _keyboardFallbackMonitor.LeftWindowsKeyReleased += OnFallbackWindowsKeyUp;
                _keyboardFallbackMonitor.Start();

                StartDeviceChangeWatcher();

                _logger.LogInfo("Waiting for Windows key press...");
                if (isInteractive)
                {
                    Console.WriteLine("\nWinToPalette is running. Press Ctrl+C to exit...");
                }

                // Start listening for key events
                await _interceptionManager.StartListeningAsync();

                // Keep the application running
                Console.CancelKeyPress += (sender, e) =>
                {
                    e.Cancel = true;
                    _logger.LogInfo("Shutting down...");
                    _interceptionManager.StopListening();
                };

                // Keep the main thread alive
                await Task.Delay(Timeout.Infinite);
            }
            catch (Exception ex)
            {
                bool isInteractive = Environment.UserInteractive && !Console.IsInputRedirected;
                
                if (_logger != null)
                {
                    _logger.LogError($"Fatal error: {ex.Message}");
                    _logger.LogError($"Stack trace: {ex.StackTrace}");
                }
                else
                {
                    Console.WriteLine($"Fatal error: {ex.Message}");
                    Console.WriteLine($"Stack trace: {ex.StackTrace}");
                }

                if (isInteractive)
                {
                    Console.WriteLine("\nPress any key to exit...");
                    WaitForKeyPress();
                }
            }
            finally
            {
                Cleanup();
            }
        }

        private static void OnWindowsKeyUp(object sender, KeyEventArgs e)
        {
            AddTimingPoint("OnWindowsKeyUp_Enter", e.Timestamp);
            TryLaunchPalette("Interception");
        }

        private static void OnFallbackWindowsKeyUp(object sender, KeyEventArgs e)
        {
            AddTimingPoint("OnFallbackWindowsKeyUp_Enter", e.Timestamp);
            TryLaunchPalette("Fallback hook");
        }

        private static void AddTimingPoint(string label, long? timestamp = null)
        {
            lock (_timingBuffer)
            {
                _timingBuffer.Add((label, timestamp ?? Stopwatch.GetTimestamp()));
            }
        }

        private static void TryLaunchPalette(string source)
        {
            bool shouldLaunch;
            lock (_launchSync)
            {
                shouldLaunch = (DateTime.UtcNow - _lastPaletteLaunchUtc).TotalMilliseconds > 400;
                if (shouldLaunch)
                {
                    _lastPaletteLaunchUtc = DateTime.UtcNow;
                }
            }

            if (!shouldLaunch)
            {
                _logger.LogDebug($"Skipping duplicate launch event ({source})");
                return;
            }

            lock (_timingBuffer)
            {
                _timingBuffer.Clear();
                _timingBuffer.Add(("TryLaunchPalette_Enter", Stopwatch.GetTimestamp()));
            }

            try
            {
                if (!_powerToysLauncher.LaunchCommandPalette(_timingBuffer))
                {
                    _logger.LogWarning($"Failed to launch PowerToys Command Palette via {source}");
                    _powerToysNotificationService?.ShowWaitingForPowerToys();
                }
                else
                {
                    FlushTimingBuffer();
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"Exception while handling Windows key release: {ex.Message}");
            }
        }

        private static void FlushTimingBuffer()
        {
            lock (_timingBuffer)
            {
                if (_timingBuffer.Count == 0)
                {
                    return;
                }

                var firstTicks = _timingBuffer[0].ticks;
                var ticksPerMs = (double)Stopwatch.Frequency / 1000.0;
                var timingLog = new System.Text.StringBuilder();
                timingLog.AppendLine("=== Launch Timing (ms from start) ===");

                foreach (var (label, ticks) in _timingBuffer)
                {
                    var elapsedMs = (ticks - firstTicks) / ticksPerMs;
                    timingLog.AppendLine($"  {elapsedMs:F3}ms - {label}");
                }

                _logger.LogDebug(timingLog.ToString());
                _timingBuffer.Clear();
            }
        }

        private static void Cleanup()
        {
            _logger?.LogInfo("Cleaning up resources...");
            StopDeviceChangeWatcher();
            _keyboardFallbackMonitor?.Dispose();
            _interceptionManager?.Dispose();
            _powerToysLauncher?.Dispose();
            _logger?.LogInfo("=== WinToPalette Launcher Stopped ===");
        }

        private static void StartDeviceChangeWatcher()
        {
            try
            {
                string query = "SELECT * FROM Win32_DeviceChangeEvent WHERE EventType = 2 OR EventType = 3";
                _deviceChangeWatcher = new ManagementEventWatcher(query);
                _deviceChangeWatcher.EventArrived += OnDeviceChanged;
                _deviceChangeWatcher.Start();
                _logger?.LogInfo("Started device change watcher");
            }
            catch (Exception ex)
            {
                _logger?.LogWarning($"Failed to start device change watcher: {ex.Message}");
            }
        }

        private static void StopDeviceChangeWatcher()
        {
            try
            {
                if (_deviceChangeWatcher != null)
                {
                    _deviceChangeWatcher.EventArrived -= OnDeviceChanged;
                    _deviceChangeWatcher.Stop();
                    _deviceChangeWatcher.Dispose();
                    _deviceChangeWatcher = null;
                }
            }
            catch (Exception ex)
            {
                _logger?.LogDebug($"Error stopping device change watcher: {ex.Message}");
            }
        }

        private static void OnDeviceChanged(object sender, EventArrivedEventArgs e)
        {
            try
            {
                if (_interceptionManager == null)
                {
                    return;
                }

                int eventType = 0;
                if (e.NewEvent.Properties["EventType"]?.Value is ushort ushortType)
                {
                    eventType = ushortType;
                }

                string reason = eventType switch
                {
                    2 => "Keyboard/device arrival detected",
                    3 => "Keyboard/device removal detected",
                    _ => "Device change detected"
                };

                _logger?.LogInfo($"{reason}. Requesting interception context recreation...");
                _interceptionManager.RequestContextRecreate(reason);

                ScheduleSelfRestartAfterDeviceChange(reason);
            }
            catch (Exception ex)
            {
                _logger?.LogDebug($"Error handling device change event: {ex.Message}");
            }
        }

        private static void ScheduleSelfRestartAfterDeviceChange(string reason)
        {
            int version = Interlocked.Increment(ref _deviceChangeVersion);

            _ = Task.Run(async () =>
            {
                await Task.Delay(1000);

                if (version != _deviceChangeVersion)
                {
                    return;
                }

                if ((DateTime.UtcNow - _lastSelfRestartUtc).TotalSeconds < 10)
                {
                    return;
                }

                _lastSelfRestartUtc = DateTime.UtcNow;
                RestartSelf($"Device-change fallback restart ({reason})");
            });
        }

        private static void RestartSelf(string reason)
        {
            try
            {
                string executablePath = Process.GetCurrentProcess().MainModule?.FileName;
                if (string.IsNullOrWhiteSpace(executablePath) || !File.Exists(executablePath))
                {
                    _logger?.LogWarning("Skipping self-restart: executable path unavailable");
                    return;
                }

                _logger?.LogWarning($"{reason}. Restarting process to rebind interception stack...");

                Process.Start(new ProcessStartInfo
                {
                    FileName = executablePath,
                    UseShellExecute = true,
                    Verb = "runas"
                });

                Environment.Exit(0);
            }
            catch (Exception ex)
            {
                _logger?.LogError($"Self-restart failed: {ex.Message}");
            }
        }

        private static void WaitForKeyPress()
        {
            try
            {
                if (Environment.UserInteractive && !Console.IsInputRedirected)
                {
                    Console.ReadKey();
                }
            }
            catch
            {
                // Ignore console access errors
            }
        }

        private static void EnsureInterceptionDll(ILogger logger)
        {
            try
            {
                string appDir = AppDomain.CurrentDomain.BaseDirectory;
                string appDllPath = Path.Combine(appDir, "interception.dll");
                
                // Check if DLL is already in app directory
                if (File.Exists(appDllPath))
                {
                    return;
                }

                // Check if DLL is in System32
                string system32Dll = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "interception.dll");
                if (File.Exists(system32Dll))
                {
                    return;
                }

                // Try to find the DLL in workspace and copy it
                string[] searchPaths = new[]
                {
                    // From build output directory up to workspace
                    Path.Combine(appDir, "..", "..", "..", "interception", "Interception_extracted", "Interception", "library", "x64", "interception.dll"),
                    Path.Combine(appDir, "..", "..", "..", "..", "interception", "Interception_extracted", "Interception", "library", "x64", "interception.dll"),
                    // Try sample directory
                    Path.Combine(appDir, "..", "..", "..", "interception", "Interception_extracted", "Interception", "samples", "x86", "interception.dll"),
                };

                foreach (string searchPath in searchPaths)
                {
                    string fullPath = Path.GetFullPath(searchPath);
                    if (File.Exists(fullPath))
                    {
                        try
                        {
                            File.Copy(fullPath, appDllPath, overwrite: true);
                            return;
                        }
                        catch (Exception ex)
                        {
                            logger.LogDebug($"Could not copy DLL: {ex.Message}");
                            // Continue to next path if copy fails
                        }
                    }
                }

                // If we reach here, the DLL wasn't found in expected locations
                // But it might still be available from System32 or other PATH locations
                logger.LogDebug($"Interception.dll not found in workspace - relying on system PATH");
            }
            catch (Exception ex)
            {
                logger.LogDebug($"Error while checking for Interception.dll: {ex.Message}");
            }
        }
    }
}
