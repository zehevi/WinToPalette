using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using WinToPalette.Logging;

namespace WinToPalette.PowerToys
{
    /// <summary>
    /// Handles launching and communication with PowerToys Command Palette
    /// </summary>
    public class PowerToysLauncher : IDisposable
    {
        private static readonly string[] CommonPowerToysLocations = new[]
        {
            // Modern per-user installation (Windows Store / MSIX)
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "PowerToys"),
            // Traditional Program Files installation
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "PowerToys"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "PowerToys"),
            // Local Programs
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "PowerToys"),
        };

        private readonly ILogger _logger;
        private readonly object _shellLock = new object();
        private Process _persistentShell;
        private StreamWriter _shellInput;
        private bool _disposed;

        public PowerToysLauncher(ILogger logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            InitializePersistentShell();
        }

        /// <summary>
        /// Initializes a persistent CMD shell for fast URI launches
        /// </summary>
        private void InitializePersistentShell()
        {
            try
            {
                _persistentShell = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = "cmd.exe",
                        UseShellExecute = false,
                        RedirectStandardInput = true,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        CreateNoWindow = true,
                        WindowStyle = ProcessWindowStyle.Hidden
                    }
                };

                _persistentShell.Start();
                _shellInput = _persistentShell.StandardInput;
                _shellInput.AutoFlush = true;

                _logger.LogInfo("Persistent shell initialized for PowerToys launches");
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Failed to initialize persistent shell: {ex.Message}");
                _persistentShell = null;
                _shellInput = null;
            }
        }

        /// <summary>
        /// Detects PowerToys installation
        /// </summary>
        public bool DetectPowerToys()
        {
            foreach (var location in CommonPowerToysLocations)
            {
                if (!Directory.Exists(location))
                {
                    continue;
                }

                var powerToysExe = Path.Combine(location, "PowerToys.exe");
                if (File.Exists(powerToysExe))
                {
                    _logger.LogInfo($"PowerToys detected at: {location}");
                    return true;
                }
            }

            _logger.LogWarning("PowerToys not detected in common locations");
            return false;
        }

        /// <summary>
        /// Launches the PowerToys Command Palette using persistent shell (fast) or fallback
        /// </summary>
        public bool LaunchCommandPalette(List<(string label, long ticks)> timingBuffer = null)
        {
            timingBuffer?.Add(("LaunchCommandPalette_Enter", Stopwatch.GetTimestamp()));

            lock (_shellLock)
            {
                timingBuffer?.Add(("Lock_Acquired", Stopwatch.GetTimestamp()));

                // Try using persistent shell first
                if (_shellInput != null && _persistentShell != null && !_persistentShell.HasExited)
                {
                    try
                    {
                        timingBuffer?.Add(("Before_WriteLine", Stopwatch.GetTimestamp()));
                        _shellInput.WriteLine("start x-cmdpal://show");
                        timingBuffer?.Add(("After_WriteLine", Stopwatch.GetTimestamp()));
                        return true;
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning($"Persistent shell failed: {ex.Message}, reinitializing...");
                        ReinitializePersistentShell();
                        
                        // Retry once with new shell
                        if (_shellInput != null)
                        {
                            try
                            {
                                timingBuffer?.Add(("Retry_WriteLine", Stopwatch.GetTimestamp()));
                                _shellInput.WriteLine("start x-cmdpal://show");
                                timingBuffer?.Add(("Retry_Complete", Stopwatch.GetTimestamp()));
                                return true;
                            }
                            catch
                            {
                                // Fall through to direct launch
                            }
                        }
                    }
                }

                timingBuffer?.Add(("Fallback_Start", Stopwatch.GetTimestamp()));

                // Fallback to direct launch if persistent shell unavailable
                try
                {
                    var process = Process.Start(new ProcessStartInfo
                    {
                        FileName = "x-cmdpal://show",
                        UseShellExecute = true
                    });

                    timingBuffer?.Add(("Fallback_ProcessStart_Complete", Stopwatch.GetTimestamp()));

                    if (process == null)
                    {
                        _logger.LogWarning("PowerToys Command Palette launch returned null process handle");
                        return false;
                    }
                    return true;
                }
                catch (Exception ex)
                {
                    _logger.LogError($"Exception while launching PowerToys: {ex.Message}");
                    return false;
                }
            }
        }

        /// <summary>
        /// Reinitializes the persistent shell after a failure
        /// </summary>
        private void ReinitializePersistentShell()
        {
            try
            {
                if (_persistentShell != null && !_persistentShell.HasExited)
                {
                    _persistentShell.Kill();
                }
            }
            catch { }

            _shellInput?.Dispose();
            _persistentShell?.Dispose();
            _shellInput = null;
            _persistentShell = null;

            // Small delay before reinitializing
            Thread.Sleep(50);
            InitializePersistentShell();
        }

        /// <summary>
        /// Checks if PowerToys Command Palette is running
        /// </summary>
        public bool IsPaletteRunning()
        {
            try
            {
                var processes = Process.GetProcessesByName("PowerToys.CommandPalette");
                return processes.Length > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error checking if palette is running: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Disposes the persistent shell and releases resources
        /// </summary>
        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            lock (_shellLock)
            {
                try
                {
                    if (_persistentShell != null && !_persistentShell.HasExited)
                    {
                        _shellInput?.WriteLine("exit");
                        if (!_persistentShell.WaitForExit(500))
                        {
                            _persistentShell.Kill();
                        }
                    }
                }
                catch { }

                _shellInput?.Dispose();
                _persistentShell?.Dispose();
                _shellInput = null;
                _persistentShell = null;
                _disposed = true;
            }
        }
    }
}
