using System;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace Desktop_Fences
{
    // Root object
    public class RemoteManifest
    {
        public RemoteMeta Meta { get; set; }
        public List<RemoteAnnouncement> Announcements { get; set; }
        public RemoteSettings Settings { get; set; }
    }

    public class RemoteMeta
    {
        public string LatestVersion { get; set; } = "0.0.0.0";
        public bool CriticalUpdate { get; set; }
        public string DownloadUrl { get; set; }
        public bool ForceKillSwitch { get; set; } // Nuclear option
    }

    public class RemoteAnnouncement
    {
        public string Id { get; set; }          // Required: Unique Key for Registry tracking
        public string Type { get; set; }        // "Info", "Warning", "Alert"
        public string Title { get; set; }
        public string Body { get; set; }
        public string Link { get; set; }        // Optional URL for "Read More"

        public string TargetVersionMin { get; set; } = "0.0.0.0";
        public string TargetVersionMax { get; set; } = "99.9.9.9";

        public int MaxDisplayCount { get; set; } = 1;   // How many times to show?
        public int AutoCloseSeconds { get; set; } = 0;  // 0 = Disable auto-close
        public bool CanUserDismiss { get; set; } = true; // Show "Don't show again" checkbox
    }

    public class RemoteSettings
    {
        public bool EnableBackgroundLogging { get; set; }
        public string SearchEngineUrl { get; set; }
    }
}