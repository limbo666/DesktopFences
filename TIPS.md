
# Desktop Fences + | Tips & Tricks

Welcome to the hidden features guide. This document highlights advanced functionality and "power-user" shortcuts to help you get the most out of Desktop Fences +.

---

## The Power of the CTRL Key
The `CTRL` key acts as a "modifier" that reveals advanced context menus and shortcuts. When in doubt, try holding `CTRL`.

* **Rename:** `CTRL + Left Click` on a fence title to rename it.
* **Arrange Icons:** `CTRL + Drag` an Icon to its new location on the fence.
* **Fence Icons Export:** `CTRL + Right Click` an empty area inside a fence to access hidden administrative tools (e.g., *Export all icons to desktop*).
* **Portal Navigation:** `CTRL + Left Click` a folder inside a **Portal Fence** to navigate into that folder within the same fence, rather than opening a new Windows Explorer window.
* **Portal Fence Naming:** `CTRL + Right Click` and select "Name Fence After Target Path" renames Portal Fence to target path.
* **Adding Seperators (spacers) Data Fences:** `CTRL + Right Click` any area on a data fence and select "Add spacer" > "Blank" or "Dot" adds a spacer into the fence. You can move it by using `CTRL + Drag` to any place you like within the fence. (v2.6.5.222 and later)
* **Export all shortcuts from a fence:** `CTRL + Right Click` any area on a data fence and select "Export all icons to desktop" 
* **Export a single shortcut:** `CTRL + Right Click` an icon and select "Send to desktop"
* **Run as different user:** `CTRL + Right Click` an icon and select "Run as a different user"  (v2.6.5.222 and later)

 ---

## Portal Fence Filters
Click the **Filter Icon** (located next to the Lock icon) to toggle the filter bar. Filters allow you to dynamically control which files are visible.

### Usage & Syntax:
* **Wildcards:** Use `*` to match patterns (e.g., `*.txt` displays only text files).
* **Multi-Filter:** Separate multiple formats with commas (e.g., `*.mp4, *.avi`).
* **Negative Filtering:** Use the `>` prefix to exclude specific terms.
    * *Example:* `*.mp3, >*live*` (Shows all MP3s EXCEPT those with "live" in the filename).
* **Persistent State:** Filters stay active until cleared. An **orange indicator** signifies an active filter. Use the **"X"** on the right of the bar to reset.


---

## Spotseach
Press spotsearch hotkey combination (default: `Ctrl + ~`) to see the search bar and search on the data fences icons (even the hidden ones)

---

## Show fences above all open windows
To quick show all fences above your open windows use `WIN + Shift + D` 

---

## Focus Fence (v2.6.5.218 and later)
Use the dedicated menu to select which fence should bring into teh front. Press hotekey combination `Ctrl + Alt + Z` (default) or use the "Focus Fence..." tray menu item.  
Double click a fence name on the list or select the name and clcik on "Focus Fence" button 


### Profiles
Profiles offering multiple configuration switching. It is a flexible method to switch icon collections between work and home, desk and presentation, work and fun or whatever mode you like to switch to.  
Profiles are the best way to keep your desktop organized and grouped according to your needs.  
Profiles can be created by using tray menu items, and/or the profile manager.   

  
Switching profiles can be done by selecting the appropriate profile name on tray menu or by using the hotkeys (default `CTRL + ALT [Profile number 0 to 9]`) or by swicthing to next (default `CTRL + ALT + .`) or previous (default `CTRL + ALT + ,`)  

  Profile Manager  
   - Profile manager can help you switch, arrange, delete and rename profiles.  

  Profile automation
   - This is a sofisticated feature that can automaticaly make Desktop Fences switch profile according to a selected program activity. Select the window you want to detect on profile automation and create the rule you want.   
   - Switch to the profile can be permanent or while the program is active.  
   - Profile automation can be turned off and on via tray menu or options window.  
  

### Smart Desktop  (v2.6.5.225 and later)
 - This is an advanced file manager for desktop which can work in the background and move files to specific folder by checking their file extension or filename.  
 - Rules can be set on "Smart Desktop Rules" window.  
 - Multiple rules can be set and notification for every successful run can be shown on the lower right screen corner to inform user that files have been moved.
 - The Smart Desktop Auto-Organize can be enabled and disbled by tray icon menu or by options window.
 - Rule priority can be set by arrnging the rule order on "Smart Desktop Rules" window.  
 

---

## Advanced JSON Tweaks
For deeper customization, you can manually edit the `options.json` file. Below are the most impactful variables:

### Filter & Display Tweaks
| Variable | Description |
| :--- | :--- |
| `NoWildcardsOnPortalFilter` | Set to `true` to match text without needing `*`. (e.g., `.mp3` instead of `*.mp3`). |
| `ShowPortalExtensions` | Set to `true` to display file extensions within Portal Fences. |
| `MaxDisplayNameLength` | Set the character limit for shortcut names (Range: `5` to `50`). |

### Workflow & Deletion
| Variable | Description |
| :--- | :--- |
| `ExportShortcutsOnFenceDeletion` | If `true`, all icons inside a fence are automatically moved to the desktop when that fence is deleted. |
| `DeleteOriginalShortcutsOnDrop` | If `true`, dropping a desktop icon into a fence will delete the original desktop file, effectively "moving" it into the fence. |

### SpotSearch
| Variable | Description |
| :--- | :--- |
| `EnableSpotSearchHotkey` | Set to `false` to completely disable the SpotSearch feature. |
| `SpotSearchKey` | Choose the trigger key: `"q"`, `"space"`, `"~"`, or a specific key code. |
| `SpotSearchModifier` | Change the modifier key to `"ALT"` or `"CONTROL"`. |
Version 2.6.5.220 and later offers UI conntrols to set hotkeys for Spotsearch so tweaking json for this is not needed anymore


### System
| Variable | Description |
| :--- | :--- |
| `DisableSingleInstance` | Set to `true` to allow running multiple separate instances of the program. |



---
