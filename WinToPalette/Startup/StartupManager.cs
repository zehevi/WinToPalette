using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using WinToPalette.Logging;

namespace WinToPalette.Startup
{
    /// <summary>
    /// Manages application startup registration using Windows Scheduled Tasks
    /// (requires admin privileges)
    /// </summary>
    public class StartupManager
    {
        private const string TaskName = "WinToPalette";
        private const string TaskPath = @"\WinToPalette\";
        private readonly ILogger _logger;

        public StartupManager(ILogger logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        /// <summary>
        /// Registers the application for startup using Windows Scheduled Tasks
        /// Requires admin privileges and should only be called during setup
        /// </summary>
        public bool RegisterForStartup(string executablePath)
        {
            try
            {
                if (!System.IO.File.Exists(executablePath))
                {
                    _logger.LogError($"Executable not found: {executablePath}");
                    return false;
                }

                // Use schtasks.exe to create the scheduled task
                // This approach is more reliable than using the COM interface directly
                var currentUser = Environment.UserDomainName + "\\" + Environment.UserName;
                var processInfo = new ProcessStartInfo
                {
                    FileName = "schtasks.exe",
                    Arguments = $"/create /tn \"{TaskPath}{TaskName}\" /tr \"\\\"{executablePath}\\\"\" /sc onlogon /ru \"{currentUser}\" /it /rl HIGHEST /f",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                };

                using (var process = Process.Start(processInfo))
                {
                    process?.WaitForExit();
                    if (process?.ExitCode == 0)
                    {
                        _logger.LogInfo($"Registered for startup via scheduled task at user logon: {executablePath}");
                        return true;
                    }
                    else
                    {
                        var error = process?.StandardError.ReadToEnd() ?? "Unknown error";
                        _logger.LogError($"Failed to create scheduled task. Exit code: {process?.ExitCode}. Error: {error}");
                        return false;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"Failed to register for startup: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Unregisters the application from startup (scheduled task)
        /// Requires admin privileges
        /// </summary>
        public bool UnregisterFromStartup()
        {
            try
            {
                var processInfo = new ProcessStartInfo
                {
                    FileName = "schtasks.exe",
                    Arguments = $"/delete /tn \"{TaskPath}{TaskName}\" /f",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                };

                using (var process = Process.Start(processInfo))
                {
                    process?.WaitForExit();
                    if (process?.ExitCode == 0 || process?.ExitCode == 1)
                    {
                        // Exit code 1 means task not found, which is fine for uninstall
                        _logger.LogInfo("Unregistered from startup");
                        return true;
                    }
                    else
                    {
                        var error = process?.StandardError.ReadToEnd() ?? "Unknown error";
                        _logger.LogError($"Failed to delete scheduled task. Exit code: {process?.ExitCode}. Error: {error}");
                        return false;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"Failed to unregister from startup: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Checks if application is registered for startup (scheduled task)
        /// </summary>
        public bool IsRegisteredForStartup()
        {
            try
            {
                var processInfo = new ProcessStartInfo
                {
                    FileName = "schtasks.exe",
                    Arguments = $"/query /tn \"{TaskPath}{TaskName}\" /fo LIST",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                };

                using (var process = Process.Start(processInfo))
                {
                    process?.WaitForExit();
                    return process?.ExitCode == 0;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error checking startup registration: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Checks if the current process is running with administrator privileges
        /// </summary>
        public static bool IsElevated()
        {
            try
            {
                var identity = System.Security.Principal.WindowsIdentity.GetCurrent();
                var principal = new System.Security.Principal.WindowsPrincipal(identity);
                return principal.IsInRole(System.Security.Principal.WindowsBuiltInRole.Administrator);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Requests elevation (re-runs application as admin)
        /// </summary>
        public static bool RequestElevation(string[] args = null)
        {
            if (IsElevated())
                return true;

            try
            {
                var processInfo = new ProcessStartInfo
                {
                    FileName = Process.GetCurrentProcess().MainModule.FileName,
                    UseShellExecute = true,
                    Verb = "runas",
                    Arguments = args != null ? string.Join(" ", args) : ""
                };

                Process.Start(processInfo);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to elevate privileges: {ex.Message}");
                return false;
            }
        }
    }
}
