using System;
using System.Drawing;
using System.Threading;
using WinToPalette.Logging;
using Forms = System.Windows.Forms;

namespace WinToPalette.Notifications
{
    public class PowerToysNotificationService
    {
        private const string WaitingTitle = "WinToPalette";
        private const string WaitingMessage = "Waiting for PowerToys to launch";
        private const int NotificationDurationMs = 2200;
        private const int NotificationCooldownMs = 5000;

        private readonly ILogger _logger;
        private readonly object _sync = new object();
        private DateTime _lastNotificationUtc = DateTime.MinValue;

        public PowerToysNotificationService(ILogger logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        public void ShowWaitingForPowerToys()
        {
            bool shouldShow;
            lock (_sync)
            {
                shouldShow = (DateTime.UtcNow - _lastNotificationUtc).TotalMilliseconds > NotificationCooldownMs;
                if (shouldShow)
                {
                    _lastNotificationUtc = DateTime.UtcNow;
                }
            }

            if (!shouldShow)
            {
                _logger.LogDebug("Skipping waiting notification due to cooldown");
                return;
            }

            var notificationThread = new Thread(() =>
            {
                try
                {
                    using var notifyIcon = new Forms.NotifyIcon
                    {
                        Icon = SystemIcons.Information,
                        Visible = true,
                        BalloonTipTitle = WaitingTitle,
                        BalloonTipText = WaitingMessage,
                        BalloonTipIcon = Forms.ToolTipIcon.Info
                    };

                    notifyIcon.ShowBalloonTip(NotificationDurationMs);
                    Thread.Sleep(NotificationDurationMs + 400);
                    notifyIcon.Visible = false;
                }
                catch (Exception ex)
                {
                    _logger.LogError($"Failed to show waiting notification: {ex.Message}");
                }
            });

            notificationThread.SetApartmentState(ApartmentState.STA);
            notificationThread.IsBackground = true;
            notificationThread.Start();
        }
    }
}
