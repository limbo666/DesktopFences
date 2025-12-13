<h1 align="center">Desktop Fences +</h1>
<p align="center"><i>Organize your desktop like magic!</i></p>
<p align="center">
<img src="https://img.shields.io/github/downloads/limbo666/DesktopFences/total?style=flat-square" alt="Total Downloads"/>
<img src="https://img.shields.io/github/stars/limbo666/DesktopFences?style=flat-square" alt="Stars"/>
  <img src="https://img.shields.io/github/forks/limbo666/DesktopFences?style=flat-square" alt="Forks"/>
  <img src="https://img.shields.io/github/issues/limbo666/DesktopFences?style=flat-square" alt="Issues"/>
  <img src="https://img.shields.io/github/license/limbo666/DesktopFences?style=flat-square" alt="License"/>
  <img src="https://img.shields.io/github/last-commit/limbo666/DesktopFences?style=flat-square" alt="Last Commit"/>
</p>

<p align="center">
  <img src="https://github.com/limbo666/DesktopFences/blob/main/Imgs/Desktop_Fences_Social_Media_PNG.png" />
  
</p>


Desktop Fences + creates **virtual fences** on your desktop, allowing you to group and organize icons in a clean and convenient way. With enhanced visual effects and right-click options, it aims to provide a more polished and customizable user interface.


**Desktop Fences+** is an open-source alternative to Stardock Fences, originally created by HakanKokcu under the name BirdyFences.

This project is a continuation and substantial modification of the original BirdyFences codebase, which was licensed under the MIT License at the time of forking. 

Desktop Fences+ has been significantly enhanced and optimized for improved performance, stability, and user experience, while respecting the terms of the original license and acknowledging the original author.

---
[![Video Title](https://img.youtube.com/vi/I-blY76kSP8/maxresdefault.jpg)](https://www.youtube.com/watch?v=I-blY76kSP8)
---

##  Manual

**A simple how to use guide is located here:** https://github.com/limbo666/DesktopFences/blob/main/desktop_fences_simple_manual.md
<br> Read it before starting an issue. 

---

##  Support

If this project has helped you, please consider supporting its development! Your contribution directly impacts how fast and far this project grows.

Maintaining and improving this tool takes time, effort, and resources. Donations help me:

 - Dedicate more time to fixing bugs and adding features

 - Cover the cost of tools (like AI assistance that speeds up development)

 - Stay motivated and keep pushing the project forward


🚀 Your support literally drives the pace of development — the more help I get, the more I can deliver!

☕ Buy me a coffee

💡 Help cover my AI subscription


Even small contributions mean a lot. Thank you for keeping this project alive and evolving! 🙏

[![Donate](https://github.com/limbo666/DesktopFences/blob/main/Imgs/donate.png)](https://www.paypal.com/donate/?hosted_button_id=M8H4M4R763RBE)

---

## Features Compared to Original BirdyFences

- **Improved Performance and Stability**
- **Fence JSON File** now placed in the same directory as the executable
- **First Fence Line** is created automatically during the first execution
- **Program Icon** added
- **Error Handlers** for Move actions, Program execution, Empty or invalid JSON files

## First release changes 🔥
- **Tray Icon** indicates the application is running
- **Program Exit Option** in the tray icon’s context menu
- **New Fence Creation** at mouse location for intuitive placement
- **About Screen**
- **Shortcuts no longer depend** on original shortcut files
- **Execution Arguments** of original shortcuts are preserved
- **Target Type Detection** with proper error handling
- **Visual Effects** on icon click and icon removal
- **Right-Click Context Menu for Icons** with: Run as Administrator (when applicable), Copy Path, Find Target
- **Folder Icon Appearance** fixed
- **Broken Link Detection** on startup and updated every second
- **Automatic JSON Format Updater** for existing fences.json
- **Fences No Longer Take Focus** from other windows when clicked
- **Options Window** to: Enable/disable snap function, Set tint level, Select base color
- **Options Saved** in `options.json`
- **Manual Backup Mechanism**: Saves fences and shortcuts to a `backups` subfolder
- **Logging Option** for diagnostics
- **Selectable Launch Effects**: Zoom, Bounce, Fadeout, SlideUp, Rotate, Agitate
- **Run at Windows Startup** option

## 2.5.0.18
- **Added**: Custom animation selection for each fence
- **Added**: Custom background color selection for each fence
- **Fixed**: 🐛 Bug on fence movement across multiscreen systems https://github.com/limbo666/DesktopFences/issues/2
- **Changed**: Start with windows option moved under `Options` window
- **Fixed**: Passing argument in `Run as administrator` selection
- **Fixed**: 🐛 Bug on portal fences causing program crash on startup when target folder is missing
- **Changed**: Code improvements on log function
- **Added**: Basic `Hide` function for each fence

## 2.5.0.23
- **Added**: ❤️ Heart menu to separate right click menu item on fences
- **Added**: Function to undo fence deletion (Restore fence)
- **Fixed**: 🐛 Bug on fence removal causing program to crash
- **Changed**: Lot of menus, descriptions, and other visual elements changed.
- **Added**: Number overlay on tray icon that indicates number of the hidden fences

## 2.5.0.26
- **Changed**: Background color codes
- **Added**: More background colors
- **Added**: More launch effects (I ♥ Elastic)
- **Fixed**: 🐛 Bug on context menus items unchecking behavior
- **Fixed**: 🐛 Bug on customization where changes were not applied into fences under same name
- **Added**: Random name generator for new fences instead of the dull "New Fence"
- **Added**: Custom delete confirmation message box

## 2.5.0.30
- **Added**: Restore for previously backed up configurations
- **Added**: Reload all Fences function
- **Added**: Tooltips on options window
- **Added**: Restore hidden fences to their hidden status on startup
- **Added**: Temporarily hide function for all fences (show desktop) on tray icon double click, restore with double click

## 2.5.1.37 - Release 3
- **Added**: Snap to Dimension function for better size alignment
- **Added**: Export Fence and Import Fence options to help move fences across systems. Few exported fences can be found [here](https://github.com/limbo666/DesktopFences/tree/main/Exported%20Fences)
- **Added**: Ability to get icons from dll libraries or executables on under `Edit` Requested on https://github.com/limbo666/DesktopFences/issues/1
- **Changed**: All message boxes changed to internal themed ones to follow mouse position across multimonitor systems
- **Changed**: Some theming correction on message boxes.
- **Added**: Sound on custom message box appearance.

## 2.5.1.40
- **Fixed**: 🐛 Bug on `Start with Windows`. The program now displays shortcuts correctly https://github.com/limbo666/DesktopFences/issues/6
- **Fixed**: 🐛 Bug with `Options` and `About` screen that misplaced controls on scaled displays https://github.com/limbo666/DesktopFences/issues/5
- **Added**: Function to display fences which are saved out of screen bounds (restored from other systems).

## 2.5.1.42
- **Fixed**: 🐛 Bug on `Portal Fences` created by misuse of FileSystemWatcher. The program now updates target files as renamed, removed. https://github.com/limbo666/DesktopFences/issues/3
- **Added**: Context menu items for `Portal Fences` and items. Now user is able to copy target file path or shortcut destination path and open `Portal Fence` target folder from right click.

## 2.5.1.58 Release 6
- **Added**: Option to show/hide tray icon (requested on https://github.com/limbo666/DesktopFences/issues/9). Attention: Hiding tray icon means you don't have access to: showing hidden fences and hiding/showing fences by double clicking on tray icon.
- **Added**: Option to show/hide portal fences watermark (requested on https://github.com/limbo666/DesktopFences/issues/11).
- **Added**: Option to use recycle bin when deleting files or folders using portal fences right click menu.
- **Added**: Lock function to fences (requested on https://github.com/limbo666/DesktopFences/issues/9).
- **Changed**: 🔨🔩🔧 Large refactoring on Log code to organize and filter logs.
- **Changed**: Options window redesigned for better user experience.
- **Fixed**: 🐛 Bug on custom message box with wrong color selection.
- **Added**: "Peek Behind" right click selection to make fences to reveal desktop contents behind them for 10 seconds.

## 2.5.1.64
- **Added**: Rollup function when `Ctrl + Click` on Fence title (📝 Use CTRL + click).
- **Added**: Function to filter hidden files on Portal Fences (request https://github.com/limbo666/DesktopFences/issues/13 and possibly fixing https://github.com/limbo666/DesktopFences/issues/14 as well).

## 2.5.1.65
- **Changed**: `Delete Fence` option moved to Heart context menu
- **Fixed**: 🐛 Bug on handling shortcuts targeting web links.
- **Added**: New icon for shortcuts targeting web links.
- **Changed**: Target check mechanism to prevent errors.
- **Added**: Indicator for network based files.

## 2.5.1.67
- **Added**: Function to re-order icons within a fence (https://github.com/limbo666/DesktopFences/issues/15). 📝 Use `CTRL + drag`.

## 2.5.1.70
- **Changed**: 🔨🔩🔧 Major code refactoring.
- **Added**: 🎇 Four new launch effects.

## 2.5.1.75
- **Added**: Error Handles on JSON loading for better stability against corrupted `fences.json` files
- **Changed**: Portal Fences are named after the target folder upon creation.
- **Changed**: Minor interface improvements for systems with resolution scaling enabled.
- **Fixed**: 🐛 Handling of shortcuts with unicode characters.
- **Changed**: Improved stability of filesystem watcher for portal fences.
- **Added**: `Rename` option for files on portal fences.
- **Added**: Lot of tweaks user can set by editing JSON files.

## 2.5.2.86
- **Changed**: Major changes to existing windows messages and other forms.
- **Changed**: 🔨🔩🔧 Another massive source code refactoring.
- **Added**: Customize window, available on fence context menu of each fence that allow user to tweak all available options.
- **Removed**: ❌ Customization submenus for effects and colors from context menu.
- **Fixed**: 🐛 Bug on Unicode folder icon.

## 2.5.2.95
- **Fixed**: 🐛 🪲 A majority of minor bugs and annoyances

## 2.5.2.111 Release 7
- **Added**: Program made single instance (with a twist 😃).
- **Added**: "Clear Dead Shortcuts" right click option to remove all not valid shortcuts from a fence.
- **Added**: "Send To Desktop" right click option to copy a shortcut to desktop (📝 Use CTRL + right click).
- **Added**: "Copy and Paste" options to copy items across fences.
- **Fixed**: 🐛 All forms were redesigned and made dpi aware to work under monitors with scaling enabled.
- **Added**: Message boxes accepted Enter and Esc as validation.
- **Added**: Variables can be set on `options.json` to manual tweak program. 🐉 See [tweaks](https://github.com/limbo666/DesktopFences/blob/main/tweaks.md).
- **Changed**: Target launch mechanism. This probably will fix most common launch errors. https://github.com/limbo666/DesktopFences/issues/24
- **Changed**: 🔨🔩🔧 Another source code refactoring.
- **Added**: Option to disable Fence scrollbars. https://github.com/limbo666/DesktopFences/issues/21
- **Fixed**: 🪲 🐛 Edit icon panel. Now arguments editing, icon preview, restore default settings bugs are eliminated.

## 2.5.2.125
- **Fixed**: 🐛 Folder name display for folder with dots on their name.
- **Fixed**: 🐛 Steam shortcuts support finally fixed https://github.com/limbo666/DesktopFences/issues/25.
- **Added**: Function to override Win + D key combination. Fences now re-appearing after Win + D is pressed https://github.com/limbo666/DesktopFences/issues/26.
- **Added**: Opacity change effect on Heart menu and Lock icons https://github.com/limbo666/DesktopFences/issues/31.
- **Added**: Option to select Menu icon and Lock icon (4 icons for each one) https://github.com/limbo666/DesktopFences/issues/31.
- **Added**: Option to set shortcut to "Always run as administrator".
- **Added**: Option to disable sound globally. 🔇 Requested on https://github.com/limbo666/DesktopFences/issues/40.
- **Added**: 🎇 Note fences (still a bit buggy but working).
- **Changed**: Double click on fence title scrolls up/down. Ctrl + Click enters rename mode. Requested on https://github.com/limbo666/DesktopFences/issues/38.

## 2.5.3.135
- **Fixed**: 🐛 Fences now are escaping the Windows Snap Assistant. https://github.com/limbo666/DesktopFences/issues/39
- **Changed**: Portal fences background image.
- **Added**: ☑️ Controls to new customization functions.
- **Changed**: Options window. Controls re-arranged. Added function to remember last tab used on options menu (during runtime).
- **Changed**: Start with Windows function changed into a more reliable approach. Now registry is used instead of startup folder. This possibly fixes the issue: 🐛 https://github.com/limbo666/DesktopFences/issues/29
- **Fixed**: 🐛 Failed to save new name bug. https://github.com/limbo666/DesktopFences/issues/45
- **Added**: Basic support for MS store based apps. This is a bare minimum implementation, no shortcut icon customization support. 🐞 Bugs expected to be found on this. https://github.com/limbo666/DesktopFences/issues/34
- **Fixed**: 🐛 Bug with program stability. Caused by icon updates and introduced during code migration.
- **Fixed**: 🐛 Bug with lost context menu on fence customization and dead shortcut cleanup. https://github.com/limbo666/DesktopFences/issues/27
- **Added**: ✨🔎📣 Search pane "SpotSearch" to search and quick launch shortcuts in all data fences. 🔥📝 Use hotkey ``CTRL+` ``
- **Changed**: 🔩 Some code refactoring.

## 2.5.3.137 Release 8
- **Added**: ☑️ Update checker engine
- **Fixed**: 🐛 Missing functions from drop icon restored.
- **Added**: ⚙️ Additional tweaks to automate icon extraction to desktop on fence deletion and icon removal from desktop on drop to fences.

## 2.5.3.140
- **Added**: ☑️ Tweak to select a key combination for SpotSearch (~, title, space, q, F1)

## 2.5.3.144
- **Changed**: Initial startup values and fences. Added a short Note fence on first start for user information.
- **Fixed**: 🐛 Bug in the fence renaming escape mechanism.
- **Added**: ☑️ Tweak to show/hide file extensions on Portal Fences.
- **Added**: ☑️ Filters for Portal Fences.

## 2.5.3.145
- **Changed**: Portal Fence filters. Added predefined filters. Added a (hidden) option to disable wildcards.
- **Added**: ☑️ Filter history for Portal Fence filters.
- **Fixed**: 🐛 Bug in tint application.

## 2.5.3.155
- **Changed**: 🎨 Customization window for Portal Fences now allows changes to icon size, icon text, etc., which were previously disabled.
- **Changed**: Snap logic improved for better Snap Near Fences functionality.
- **Changed**: 🆒 Right-click menu items on Data Fences are now populated dynamically and displayed only when actions are available.
- **Added**: ☑️ Ability for Portal Fences to display folders with their original customized icons.
- **Added**: ☑️ Daily automatic backup functionality.
- **Added**: ☑️ Cut and Copy options in the right-click menu for Portal Fence items.
- **Added**: ☑️ Reset and Clear All Data buttons in Options.

## 2.5.4.161
- **Added**: 💥 Tabs❕ 👋 Half of the tabs code was already present, but this release involved a long journey of debugging, stabilizing, and improving functionality to finally bring tabs to life.
- **Fixed**: 🐛 Minor bugs discovered during tab development. Overall program stability improved.
- **Changed**: Backup engine updated. The program now backs up and restores options.json (if selected during restore).

## 2.5.4.170
- **Added**: 🎨 Icon customization option for shortcuts targeting folders.
- **Added**: Support for rearranging icons on tabs using CTRL + Drag.

## 2.5.4.172
- **Fixed**: 💊 Visual glitches in various program functions.
- **Fixed**: 🐛 Missing context menu for weblinks dropped into fences.
- **Added**: 🎨 Icon customization option for shortcuts targeting weblinks. (This one was not easy.)

## 2.5.4.178
- **Changed**: Massive tab-related adjustments.
- **Changed**: Backup and Export/Import engine updated to support tabs.
- **Fixed**: 🐛 Several minor bugs.

## 2.5.4.183
- **Changed**: Heart Menu rearranged.
- **Added**: Tab overflow management and navigation mechanism.
- **Fixed**: 🐛 Bug in settings application from the Options window.

## 2.5.4.184
- **Changed**: Log engine optimized to reduce file system load.

## 2.5.4.186
- **Added**: Basic folder navigation for Portal Fences.
- **Changed**: Tab naming pattern.

---

##  Summary

Desktop Fences + brings a powerful and visually optimized experience for users who want to organize their desktop with flexibility and style. The program is designed to enhance productivity by combining intuitive interactions, aesthetic customization, and practical right-click options.

---

##  Release
Get the latest release here:
https://github.com/limbo666/DesktopFences/releases


---

##  Installation

Simply extract the executable and run it. The necessary configuration files (`fences.json`, `options.json`) will be created on first run.
Attention: Only user-writable locations are compatible. See (https://github.com/limbo666/DesktopFences/issues/51)

>  Compatible with Windows 10/11  
>  No installation required — fully portable

---
## Stargazers over time
[![Stargazers over time](https://starchart.cc/limbo666/DesktopFences.svg?background=%23535050&axis=%23f172cb&line=%239c34f7)](https://starchart.cc/limbo666/DesktopFences)
---

##  License

This project is licensed under the [MIT License](License.md).

---

##  Credits

Based on the original **BirdyFences** by HakanKokcu  
Desktop Fences + is Enhanced and maintained by Nikos Georgousis. 
Hand Water Pump 2025









