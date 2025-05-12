using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;

namespace Desktop_Fences
{
    /// <summary>
    /// Manages backup operations for Desktop Fences data, including the fences configuration file
    /// and associated shortcut files. Backups are stored in timestamped folders for easy recovery.
    /// </summary>
    public static class BackupManager
    {



        public static void BackupData()
        {
            try
            {
                // Get the directory of the executing assembly
                string exePath = Assembly.GetEntryAssembly().Location;
                string exeDir = Path.GetDirectoryName(exePath);

                // Define source paths for fences.json and Shortcuts folder
                string jsonFilePath = Path.Combine(exeDir, "fences.json");
                string shortcutsFolderPath = Path.Combine(exeDir, "Shortcuts");

                // Define the destination "Backups" folder
                string backupsFolderPath = Path.Combine(exeDir, "Backups");
                if (!Directory.Exists(backupsFolderPath))
                {
                    Directory.CreateDirectory(backupsFolderPath); // Create Backups folder if it doesn't exist
                }

                // Create a new backup folder with the current date and time (e.g., "2503181234_backup")
                string backupFolderName = DateTime.Now.ToString("yyMMddHHmm") + "_backup";
                string backupFolderPath = Path.Combine(backupsFolderPath, backupFolderName);
                Directory.CreateDirectory(backupFolderPath);

                // Copy the fences.json file to the backup folder
                string backupJsonFilePath = Path.Combine(backupFolderPath, "fences.json");
                System.IO.File.Copy(jsonFilePath, backupJsonFilePath, true); // Overwrite if exists

                // Copy the entire "Shortcuts" folder, if it exists
                string backupShortcutsFolderPath = Path.Combine(backupFolderPath, "Shortcuts");
                if (Directory.Exists(shortcutsFolderPath))
                {
                    Directory.CreateDirectory(backupShortcutsFolderPath);

                    // Recursively copy all subdirectories
                    foreach (string dirPath in Directory.GetDirectories(shortcutsFolderPath, "*", SearchOption.AllDirectories))
                    {
                        Directory.CreateDirectory(dirPath.Replace(shortcutsFolderPath, backupShortcutsFolderPath));
                    }

                    // Copy all files
                    foreach (string filePath in Directory.GetFiles(shortcutsFolderPath, "*.*", SearchOption.AllDirectories))
                    {
                        string destFilePath = filePath.Replace(shortcutsFolderPath, backupShortcutsFolderPath);
                        System.IO.File.Copy(filePath, destFilePath, true); // Overwrite if exists
                    }
                }

                // Notify the user of successful backup
                //  MessageBox.Show("Backup completed successfully.", "Backup", MessageBoxButton.OK, MessageBoxImage.Information);
                TrayManager.Instance.ShowOKOnlyMessageBoxForm("Backup completed successfully.", "Backup");

            }
            catch (Exception ex)
            {
                // Handle any errors during the backup process and inform the user
                //MessageBox.Show($"An error occurred during backup: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                TrayManager.Instance.ShowOKOnlyMessageBoxForm($"An error occurred during backup: {ex.Message}", "Error");

            }
        }
    }
}