# Desktop Fences +

<p align="center">
  <img src="https://github.com/user-attachments/assets/fb7f3da9-2e1f-4f79-90dc-33994ff33104" />
</p>


Desktop Fences + creates **virtual fences** on your desktop, allowing you to group and organize icons in a clean and convenient way. With enhanced visual effects and right-click options, it aims to provide a more polished and customizable user interface.

**Desktop Fences +** is an open-source alternative to **Stardock Fences**, originally created by [HakanKokcu](https://github.com/HakanKokcu) under the name **BirdyFences**.

**Desktop Fences +** has been significantly enhanced and optimized for improved performance, stability, and user experience.

---

##  Support
If you find this project helpful, consider supporting its development!
Your contribution helps cover development time, new features, bug fixes, and overall project improvements.

 - Buy me a coffee

 - Help me pay AI subscription (yeah, development gets easier with AI but costs money)

Every little bit helps and is greatly appreciated! 🙏


[![Donate](https://github.com/limbo666/DesktopFences/blob/main/Imgs/donate.png)](https://www.paypal.com/donate/?hosted_button_id=M8H4M4R763RBE)

---

##  Features Compared to Original BirdyFences

-  **Improved Performance and Stability**
-  **Fence JSON File** now placed in the same directory as the executable
-  **First Fence Line** is created automatically during the first execution
-  **Program Icon** added
-  **Error Handling** on: Move actions, Program execution, Empty or invalid JSON files

##  New Functionalities
-  **Minimal About Screen**
-  **Tray Icon** indicates the application is running
-  **Program Exit Option** in the tray icon’s context menu
-  **New Fence Creation** at mouse location for intuitive placement
-  **Shortcuts no longer depend** on original shortcut files
-  **Execution Arguments** of original shortcuts are preserved
-  **Target Type Detection** with proper error handling
-  **Visual Effects** on icon click and icon removal
-  **Right-Click Context Menu for Icons** with: Run as Administrator (when applicable), Copy Path, Find Target
-  **Folder Icon Appearance** fixed
-  **Broken Link Detection** on startup and updated every second
-  **Automatic JSON Format Updater** for existing fences.json
-  **Fences No Longer Take Focus** from other windows when clicked
-  **Options Window** to: Enable/disable snap function, Set tint level, Select base color
-  **Options Saved** in `options.json`
-  **Manual Backup Mechanism**:  Saves fences and shortcuts to a `backups` subfolder
-  **Logging Option** for diagnostics
-  **Selectable Launch Effects**: Zoom, Bounce, Fadeout, SlideUp, Rotate, Agitate
-  **Run at Windows Startup** option

##  2.5.0.18
-  **Added**:  Custom animation selection for each fence
-  **Added**:  Custom background color selection for each fence
-  **Fixed**:  Bug on fence movement across multiscreen ssystem
-  **Changed**:  Start with windows option moved under `Options` window
-  **Fixed**: Passing argument in `Run as administrator` selection
-  **Fixed**: Bug on portal fences causing program crash on startup when target fodler is missing
-  **Changed**:  Code improvements on log function
-  **Added**:  Basic `Hide` function for each fence
  

##  2.5.0.23
-  **Added**:  Heart menu to seperate rightclick menu item on fences 
-  **Added**:  Function to undo fence deletion (Restore fence)
-  **Fixed**:  Bug on fence removal causing program to crash
-  **Changed**:  Lot of menus, descriptions, and other visual elements changed.
-  **Added**:  Number overlay on tray icon that indicates number of the hidden fences 


##  2.5.0.26
-  **Changed**:  Background color codes
-  **Added**:  More background colors
-  **Added**:  More launch effects (I ♥ Elastic)
-  **Fixed**:  Bug on context menus items unchecking behavior
-  **Fixed**:  Bug on customization where chenges was not applied into fences with same name
-  **Added**:  Random name generator for new fences instead of the dull "New Fence"
-   
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

>  Compatible with Windows 10/11  
>  No installation required — fully portable

---

##  License

This project is licensed under the [MIT License](License.md).

---

##  Credits

Based on the original **BirdyFences** by [HakanKokcu](https://github.com/HakanKokcu)  
Enhanced and maintained by the Desktop Fences + community.









