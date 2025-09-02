# Desktop Fences + User Manual


<div style="width:20%; margin: auto;">

![blackboard fences](https://github.com/limbo666/DesktopFences/blob/main/Imgs/Fence2.png) 
</div>


## Overview

Desktop Fences + is a powerful desktop organization tool that creates virtual "fences" on your desktop, allowing you to group and organize icons in a clean and convenient way. Think of fences as visible containers that help you organize your desktop shortcuts, files, and folders into logical groups.

---

## Getting Started

### System Requirements

- Windows 10 or Windows 11 operating system
- .NET Framework support
- Administrative privileges for some features

### First Launch

When you first start Desktop Fences +, the application will:

1. Create necessary basic configuration files
2. Add itself to the system tray

---

## Core Features

### 1. Creating Fences

**What are Fences?**
Fences are containers on your desktop that group related shortcuts and files together. You can think of them as boxes that keep your desktop organized.

**Fence Types:**

- **Fences (Data Fences)**: Regular containers for shortcuts targeting files or folders and web links
- **Portal Fences**: Special fences mirroring the contents of folders

**How to Create a Fence:**

1. The first fence is created automatically on first run. look on the top left corner of your screen
2. From there clicking the heart menu you can select New Fence to create a new Fence or New Portal Fence to create a new Portal Fence

**Rename a fence:**

- Double click the fence title. Type your preferred name and click out of the edit area to finish.

### 2. Adding Items to Fences

**Adding Shortcuts:**

- Drag and drop existing desktop shortcuts, folders, or files into any fence. This will generate automatically a shortcut for the target.

  **Supported types**

- Shortcuts (.lnk files)
- Web links (.url files), shortcuts targeting web sites or links from web browsers
- Folders and directories
- Document files
- Network targers

### 3. Moving Items Between Fences

**Using the Move Dialog:**

1. Right-click on any item in a fence
2. Select "Move to..." from the context menu
3. Choose the destination fence from the hierarchical list
4. Confirm the move operation

**Using Copy Item:**

1. Right-click on any item in a fence
2. Select "Copy Item" from the context menu
3. Right-click on any other Fence and select "Paste Item"

---

## Customization Options

### 1. Fence Appearance

Access fence customization by right-clicking on a fence title bar and selecting "Customize...":
This will show the Customize Fence window.



**Fence Settings:**

- **Custom Color**: Choose from 13 preset colors (Red, Green, Teal, Blue, Bismark, White, Beige, Gray, Black, Purple, Fuchsia, Yellow, Orange)
- **Custom Launch Effect**: Customize the effect on icon click. See Launch Effects below
- **Border Color**: Customize the fence border appearance
- **Border Thickness**: Adjust border width (1-5 pixels)

**Title**

- **Title Text Color**: Change fence title color
- **Title Text Size**: Small, Medium, or Large
- **Bold Title Text**: Make fence titles bold

**Icons**

- **Icon Size**: Tiny (24px), Small (32px), Medium (40px), Large (48px), Huge (64px)
- **Icon Spacing**: Adjust space between icons (0-20 pixels)
- **Text Color**: Color for item labels
- **Disable Text Shadow**: Remove shadow effects from text
- **Grayscale Icons**: Convert all icons to grayscale ()

### 2. Launch Effects

When you click on items in fences, choose from these visual effects:

- **Zoom**: Items grow before launching
- **Bounce**: Items bounce up and down
- **FadeOut**: Items fade away
- **SlideUp**: Items slide upward
- **Rotate**: Items spin before launching
- **Agitate**: Items shake vigorously
- **GrowAndFly**: Items grow and fly off screen
- **Pulse**: Items pulse with light
- **Elastic**: Items stretch elastically
- **Flip3D**: Items flip in 3D space
- **Spiral**: Items spiral outward
- **Shockwave**: Creates a shockwave effect
- **Matrix**: Matrix-style digital rain effect
- **Supernova**: Explosive light effect
- **Teleport**: Items teleport with particle effects


## Change Icon Order

Hold down the CTRL button and drag the icon to its new position within the fence.


## Customize Shortcuts

You can check and change the icons properties by right clicking the icon and selecting "Edit..."

In the edit window you can set name, target path, arguments and icon for the shortcut. Note: The program supports ico, exe and dll files as icon sources. 



## Manage Invalid Shortcuts

The icons on the fences which are targeting files and folders are getting continuously checked for target validity. If target is missing or it is not accessible an icon indicate that this shortcut is not functioning, so you can check the target and fix it or delete the dead shortcut. You can right click on this icon and select "Remove".   

**Clear Dead Shortcuts :** If you want to remove all dead shortcuts within a fence with just one move right click on the fence and select "Clear Dead Shortcuts"



## Convenient Usage Functions

**Hide Fence**: Right click a fence an select "Hide Fence". This sets the fence as hidden and the number of hidden fences increases ion the tray icon. To show the fence again you can use the tray icon context menu where the hidden fences are shown as items on "Show Hidden Fences" menu

**Peek Behind**: Right click a fence an select "Hide Fence". The fence hides for 10 second to help access the desktop behind the fence. This feature is helpful if for some reason you need to see behind the fence for a short time. A count down timer is shown in place.

**Hide/Unhide All Fences **: Double click the tray icon. All fences will toggle. Note: This function doesn't change the status of hidden fences with teh above mentioned "Hide Fence" function.

**Rollup/RollDown a fence** : Ctrl + Click on fence title. The fence will rollup into its title. To Roll down again to fence previous height Ctrl + Click again the title area.

**Fence Lock :** Click on the shield icon🛡️on top right corner of the fence. This locks fence position and size. Rollup/Rolldown is still available when locked.



## Advanced usage of shortcuts

Right clicking any shortcut the following usage options are available:
**Run as Administrator :** Executes the target program using administrator rights. Note: This uses also the passed arguments on the icon.

**Copy Path -> Folder Path** : Copies the target folder path

**Copy Path -> Full Path** : Copies the target file path

**Open target folder** : Open target folder in file explorer





## Import/Export fences

Click on the heart ❤️ menu and select "Export this Fence" to export a fence into a  *.fence file.  This file will be exported into "Exports" subfolder on program path.
To import a fence exported previously click on the heart ❤️ menu and select "Import a Fence", browse for a *.fence file and import it.



---

## Application Settings

Access settings through the system tray icon → "Options":

### General Tab

**Startup:**

- **Start with Windows**: Automatically launch with system startup

  

**Selections:**

- **Single Click to Launch:** Enables single click on icons, otherwise double click is needed

- **Enable Snap Near Fences**: Automatically align fences to each other

- **Enable Dimension Snap**: Automatically snaps size to the closest value of multiple of 10

- **Enable Tray Icon**: Display icon in system notification area

- **Enable Portal Fences Watermark:**  Enables the "portal" watermark image on portal fences

- **Use Recycle Bin on Portal Fences 'Delete Item' command:** Send files to recycle bin instead of deleting them

  

**Style:**

- **Tint **: Adjust overall fences tint (0-100%)
- **Color**: Sets the default fences color
- **Launch Effect**: Sets the default launch effect

### Tools Tab

**Backup and Restore:**

- **Backup Data**: Create timestamped backups of all fences and shortcuts
- **Import Backup**: Restore from previous backup files
- **Open Backups Folder:** Opens backups folder in file explorer



### Look Deeper Tab

- **Enable Logging **: Enable detailed logging

- **Open Log **: Open log file for viewing

  

---



---

## Troubleshooting

### Common Issues

**Elements on help forms are not aligned properly **

1. Check the resolution scaling.

2. Set scaling to 100% if it is possible.

   This is an ongoing development and will be fixed in future versions 

**Performance Issues**

1. Disable logging if enabled. Log can delay the program performance

2. Set launch effect to Zoom or FadeOut

   

**Configuration Issues**

1. Reset the program! Delete options.json and fences.json from progarm fodler and try again



**Portal Fences Issues**

1. Delete portal fence from fences.json and try creating again
2. Verify the path exists and it is accessible



---

## Tips for Best Use

### Organization Strategies

**By Category:**

- Create fences for different types of applications (Games, Work, Utilities)
- Use color coding to visually distinguish categories
- Keep related items together for easy access

**By Frequency:**

- Place frequently used items in easily accessible fences
- Use larger icon sizes for important applications
- Consider using launch effects for visual feedback

**By Project:**

- Create temporary fences for specific projects

- Include all related files, documents, and tools

  

---

## Technical Information

### File Locations

**Configuration Files:**

- `fences.json`: Main fence configuration and layout
- `Desktop_Fences.log`: Application log file
- `Shortcuts/`: Folder containing fence shortcut files

**Backup Structure:**

- `Backups/[TIMESTAMP]_backup/`: Timestamped backup folders
- Contains both configuration and shortcut files
- Compressed archives for easy storage and transfer

**Dependencies:**

- .NET Framework
- Windows Forms and WPF libraries
- Icon extraction libraries
- JSON processing components

---



### Version Information

The application version is displayed in the About dialog, accessible through the system tray menu. This information is useful when reporting issues or checking for updates.

---

*This manual covers the core functionality of Desktop Fences +. The application includes many additional features and customization options discoverable through exploration and experimentation.*

*Ctrl + Click is your friend*

Aug 2025 Nikos Georgousis
Hand Water Pump

