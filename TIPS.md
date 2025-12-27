
# 💡 Desktop Fences + | Tips & Tricks

Welcome to the hidden features guide. This document highlights advanced functionality and "power-user" shortcuts to help you get the most out of Desktop Fences +.

---

## ⌨️ The Power of the CTRL Key
The `CTRL` key acts as a "modifier" that reveals advanced context menus and shortcuts. When in doubt, try holding `CTRL`.

* **Quick Rename:** `CTRL + Left Click` on a fence title to rename it instantly.
* **Hidden Icon Options:** `CTRL + Right Click` an icon to see advanced actions (e.g., *Send to Desktop*).
* **Fence Management:** `CTRL + Right Click` an empty area inside a fence to access hidden administrative tools (e.g., *Export all icons to desktop*).
* **Portal Navigation:** `CTRL + Left Click` a folder inside a **Portal Fence** to navigate into that folder within the same fence, rather than opening a new Windows Explorer window.

---

## 🔍 Portal Fence Filters
Click the **Filter Icon** (located next to the Lock icon) to toggle the filter bar. Filters allow you to dynamically control which files are visible.

### Usage & Syntax:
* **Wildcards:** Use `*` to match patterns (e.g., `*.txt` displays only text files).
* **Multi-Filter:** Separate multiple formats with commas (e.g., `*.mp4, *.avi`).
* **Negative Filtering:** Use the `>` prefix to exclude specific terms.
    * *Example:* `*.mp3, >*live*` (Shows all MP3s EXCEPT those with "live" in the filename).
* **Persistent State:** Filters stay active until cleared. An **orange indicator** signifies an active filter. Use the **"X"** on the right of the bar to reset.

---

## 🛠️ Advanced JSON Tweaks
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

### SpotSearch & System
| Variable | Description |
| :--- | :--- |
| `EnableSpotSearchHotkey` | Set to `false` to completely disable the SpotSearch feature. |
| `SpotSearchKey` | Choose the trigger key: `"q"`, `"space"`, `"~"`, or a specific key code. |
| `SpotSearchModifier` | Change the modifier key to `"ALT"` or `"CONTROL"`. |
| `DisableSingleInstance` | Set to `true` to allow running multiple separate instances of the program. |

---
