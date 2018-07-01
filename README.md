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

1. Open your .brd file in eagle
2. From the ULP menu run the "eagleJSONexporter_v1_0.ulp". You should see the following dialog
![alt text](https://github.com/lagnajeet/ECADWorksAddin/blob/master/images/MCAD.png "Eagle JSON export ULP")
3. Set the board thickness (in mm) and click "Export JSON". Save the generated JSON file somewhere in your computer.
4. In Solidworks Click on "Open Board" and browse for the JSON file you just created.
5. If all goes well you should see the board with cutouts and holes and the component list dropdown popolated with the components in the board
6. Select a style preset and click "Apply style". It will make the traces, silkscreen, pads etc on the board. It makes it easy to place components (in case they don't go to the right place automatically). The preset style drop down list has some predefined styles. It shows the colors on the left side of the style name in the list for quick preview of the color pellete of the style. More styles can be added to the styles file located at %appdata%\ECADWorks\PCBStyle.json. The colors in the style fil follow the ARGB 8 bit format.
7. Select a component from the component list under the "Component settings" group box. The "assigned file" and "number of instances"  will be updated. If the component has no solidworks model (i.e. a sldprt file) assigned to it then it will have a red cross next to it name in the list and the "Assigned File" field will say "None"
8. You can check the boxes Top and bottom next o the field "Highlight select on". It will show the place where the current selected model goes on the board.
9. Click "Browse Model" and select the solidworks model for the component. It will place the model on the board. If for some reason you need to remove the model then just click "Clear model". It will un assign the model and clear it from the board.
10. The program remembers all the past component to model assignments and build a library. The library is stored as s json file in the appdata folder. (%appdata%\ECADWorks\PackageLibrary.json). Clicking on "Load Library" will load models for components present in the library file. Checking "Load Selected" will load library (if any exist) for the current selected component in the component drop down list.
11. Most of the model for various electrical/electronics/optical components can be found at [3dcontentcentral](https://3dcontentcentral.com). But as they are all third party made models they may not be following the same axis convention as the board assembly used in this addin. In which case you might have to rotate/translate it to make it fit on the board properly.
12. That brings us to the "Transformations" section. Here you can rotate and translate the components. The transformations apply to all instances of the model for the selected component. If you want to apply transformation to a particullar model select it in the right side view and check the box "Transform only the selected ones". The "Auto Align" button will align and place the model on the board. It does this by calculating the centrot of the model. It may not be always accurate and might need some manual fiddling. If the package is an SMD package check the box labled "SMD Package" uncheck it if it a through hole package.
13. You can save the board file by clicking "Save Board" it will save it as a JSON file that when opened through the ECADWorks Add-in will generate the board and the coponents with the corresponding transformations.
14. The addin automatically saves (on file close event) the location of the model assigned to each component and the corresponding transformations (if any) to the library file. That way the model can be loaded from the library in future if needed.
