# Eagle to Solidworks 3D Addin

*This is an experimental program at this point. Please keep that in mind while using the ULP and/or the Add-in.*

This solidworks Add-in combined with the Eagle ulp does the following.
### Converts
![alt text](https://github.com/lagnajeet/ECADWorksAddin/blob/master/images/ECAD.gif "Eagle CAD file")

### Into
![alt text](https://github.com/lagnajeet/ECADWorksAddin/blob/master/images/MCAD.png "Solidworks 3D render")

### How to install

1. Install the ULP
- Download the "eagleJSONexporter_v1_0.zip" file.
- Extract/copy the "eagleJSONexporter_v1_0.ulp" to the ulp directory of eagle installation.

2. Install the Solidworks Add-in
- If you are only interested in installing the component and not compiling it on your own then all you need is "ECADWorksAddin-Release.zip"
- You also need the Add-in installed by [angelsix](https://github.com/angelsix) "SolidWorksAddinInstaller.exe". It simplifies installing solidworks Add-ins. It's quite simple. It should populate the path of Solidworks and RegAsm automatically. If your solidworks is not installed at the default location then just browse for solidworks.exe file for the first field. He has a [nice video](https://youtu.be/7DlG6OQeJP0?t=2373) explaing how the installer works 
- Extract the containts of "ECADWorksAddin-Release.zip" somewhere in you computer. It does't matter where. If not sure extract the folder insdie the zip file to solidworks installation folder.
- Use "SolidWorksAddinInstaller.exe" to install the dll file "LP.SolidWorks.ECADWorksAddin.dll".
- At this point open SolidWorks and open the Add-in dialog box. (Click on the down arrow near the settings icon) 

![alt text](https://github.com/lagnajeet/ECADWorksAddin/blob/master/images/enable.gif "Solidworks Add-in dialog")

- Check the box under "Start Up" if you want to load this Add-in automatically everytime solidworks starts. Click Ok. And you should get the following panel on the left side.

![alt text](https://github.com/lagnajeet/ECADWorksAddin/blob/master/images/ECADWorks.gif "Solidworks ECADWorks Dialog") 

You might have to click on the Icon to open the settings box.

### How to use
