using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using System.Reflection;
using Newtonsoft.Json;

namespace Desktop_Fences
{
    public static class RemoteInfoManager
    {
        // YOUR REAL GITHUB URL
        private const string MANIFEST_URL = "https://raw.githubusercontent.com/limbo666/DesktopFences/refs/heads/main/ngdfcs/getversion.json";

        private static bool _hasChecked = false;

        public static void Initialize()
        {
            if (_hasChecked) return;
            _hasChecked = true;

            // Run in background
            Task.Run(async () =>
            {
                // Wait 25 seconds so we don't slow down startup
                await Task.Delay(25000);
                await PerformCheck();
            });
        }

        private static async Task PerformCheck()
        {
            try
            {
                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.UserAgent.ParseAdd("DesktopFences/2.5"); // Identify app
                    client.Timeout = TimeSpan.FromSeconds(10);

                    // FIX: Add Cache Buster to force fresh content from GitHub
                    // We append a unique timestamp so the CDN/Proxy treats this as a new request.
                    string urlWithNoCache = $"{MANIFEST_URL}?t={DateTime.Now.Ticks}";

                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General, $"RemoteInfoManager: Downloading manifest from {urlWithNoCache}");

                    string json = await client.GetStringAsync(urlWithNoCache);

                    if (string.IsNullOrWhiteSpace(json)) return;

                    var manifest = JsonConvert.DeserializeObject<RemoteManifest>(json);
                    if (manifest == null) return;

                    // 1. Process Updates & Killswitch
                    ProcessMeta(manifest.Meta);

                    // 2. Process Settings (Silent)
                    if (manifest.Settings != null)
                    {
                        if (manifest.Settings.EnableBackgroundLogging)
                            SettingsManager.EnableBackgroundValidationLogging = true;
                    }

                    // 3. Process Announcements (UI)
                    if (manifest.Announcements != null && manifest.Announcements.Count > 0)
                    {
                        foreach (var msg in manifest.Announcements)
                        {
                            if (ShouldShowMessage(msg))
                            {
                                System.Windows.Application.Current.Dispatcher.Invoke(() =>
                                {
                                    // SHOW THE NOTIFICATION WINDOW
                                    NotificationFormManager.Show(msg);
                                });
                                break; // Show only one announcement per session
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Fail silently on network errors
                LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.General, $"RemoteInfoManager Error: {ex.Message}");
            }
        }

        private static void ProcessMeta(RemoteMeta meta)
        {
            if (meta == null) return;

            if (meta.ForceKillSwitch)
            {
                System.Windows.Application.Current.Dispatcher.Invoke(() => System.Windows.Application.Current.Shutdown());
                return;
            }

            // VERSION CHECK (4-Part Support: Major.Minor.Build.Revision)
            Version currentVer = Assembly.GetEntryAssembly().GetName().Version;

            if (Version.TryParse(meta.LatestVersion, out Version remoteVer))
            {
                // This comparison checks version
                if (remoteVer > currentVer)
                {
                    // Construct a fake "Announcement" for the update so we can reuse the UI
                    var updateMsg = new RemoteAnnouncement
                    {
                        Id = $"UPDATE_{remoteVer}", // Unique ID ensures they see it once/max count
                        Title = "Update Available",
                        Body = $"A new version ({remoteVer}) is available.\nYou are currently using {currentVer}.",
                        Type = meta.CriticalUpdate ? "Alert" : "Info",
                        Link = !string.IsNullOrEmpty(meta.DownloadUrl) ? meta.DownloadUrl : "https://github.com/limbo666/DesktopFences/releases",
                        CanUserDismiss = !meta.CriticalUpdate,
                        MaxDisplayCount = meta.CriticalUpdate ? 50 : 3 // Nag more for critical
                    };

                    if (ShouldShowMessage(updateMsg))
                    {
                        System.Windows.Application.Current.Dispatcher.Invoke(() =>
                        {
                            NotificationFormManager.Show(updateMsg);
                        });
                    }
                }
            }
        }

        private static bool ShouldShowMessage(RemoteAnnouncement msg)
        {
            if (string.IsNullOrEmpty(msg.Id)) return false;

            // 1. Check Dismissal
            if (RegistryHelper.IsMessageDismissed(msg.Id)) return false;

            // 2. Check Max Count
            int currentCount = RegistryHelper.GetMessageDisplayCount(msg.Id);
            if (currentCount >= msg.MaxDisplayCount) return false;

            // 3. Check Version Targeting
            Version currentVer = Assembly.GetEntryAssembly().GetName().Version;

            if (Version.TryParse(msg.TargetVersionMin, out Version minVer) && currentVer < minVer) return false;
            if (Version.TryParse(msg.TargetVersionMax, out Version maxVer) && currentVer > maxVer) return false;

            return true;
        }
    }
}