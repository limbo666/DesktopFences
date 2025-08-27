using IWshRuntimeLibrary;

namespace Desktop_Fences
{
    public class ShortcutManager
    {
        //public static void CreateShortcut(string sourceShortcutPath, string destinationPath)
        //{
        //    WshShell shell = new WshShell();
        //    IWshShortcut sourceShortcut = (IWshShortcut)shell.CreateShortcut(sourceShortcutPath);

        //    string targetPath = sourceShortcut.TargetPath;
        //    string arguments = sourceShortcut.Arguments;
        //    string iconLocation = sourceShortcut.IconLocation;

        //    IWshShortcut newShortcut = (IWshShortcut)shell.CreateShortcut(destinationPath);
        //    newShortcut.TargetPath = targetPath;
        //    newShortcut.Arguments = arguments;
        //    newShortcut.IconLocation = iconLocation;
        //    newShortcut.Save();
        //}

        //public static void EditShortcut(string shortcutPath, string newTargetPath)
        //{
        //    WshShell shell = new WshShell();
        //    IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);

        //    string iconLocation = shortcut.IconLocation;

        //    shortcut.TargetPath = newTargetPath;
        //    shortcut.IconLocation = iconLocation;
        //    shortcut.Save();
        //}
    }
}