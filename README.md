RallyRestToolkitForVBA
======================

##What Is It?

A Rally REST toolkit written for Microsoft Visual Basic for Applications (VBA). VBA is a scripting language used by Microsoft Office (and other) applications.

##Why?
Why not? VBA is a useful toolkit for automating processes in Microsoft Excel, for one. The [Rally Excel Plugin](https://help.rallydev.com/rally-add-excel "Rally Excel Add-In"), while providing a nice UI-based interface for querying Rally, does not offer script-ability or automation.

##Why Not?
Excellent question. VBA may be quite slow and inefficient for accessing large amounts of data. It's not asynchronous, and can be a clumsy way to get data into/out of Rally in large volumes. But for small datasets, it could be convenient.

##How functional is this toolkit at this point in time?
Alpha-level code. Right now this is basically a proof-of-concept for Querying Rally and Creating/Updating Rally Artifacts from VBA code. More to come though!

## Getting started

*Note: You'll need to follow ALL of the steps below before trying any of the example code.*

1. Download the [RallyRestToolkitForVBA.xlsm](https://github.com/markwilliams970/RallyRestToolkitForVBA/blob/master/ExcelWorksheet/RallyRestToolkitForVBA.xlsm?raw=true "RallyRestToolkitForVBA.xlsm") Excel Worksheet

2. Enable Macros (Enable Content)

![Enable Macros](https://raw.githubusercontent.com/markwilliams970/RallyRestToolkitForVBA/master/screenshots/screenshot1.png)

3. Show the Developer Tools Menu in Excel
4. File -> Options ->

![File Options](https://raw.githubusercontent.com/markwilliams970/RallyRestToolkitForVBA/master/screenshots/screenshot2.png)

5. Customize Ribbon -> Developer "Checked"

![Customize Ribbon / Developer](https://raw.githubusercontent.com/markwilliams970/RallyRestToolkitForVBA/master/screenshots/screenshot3.png)

5. Open Visual Basic Editor

![VBA Editor](https://raw.githubusercontent.com/markwilliams970/RallyRestToolkitForVBA/master/screenshots/screenshot4.png)

6. Go to Tools -> References

![Tools References](https://raw.githubusercontent.com/markwilliams970/RallyRestToolkitForVBA/master/screenshots/screenshot5.png)

7. The References shown here are needed to use the RallyRestToolkitForVBA. Add/load any that are not checked on in your environment.

![Needed References](https://raw.githubusercontent.com/markwilliams970/RallyRestToolkitForVBA/master/screenshots/screenshot6.png)

## Example Video

[Video showing Sample Query](http://www.screencast.com/t/OGxiqMAXxi5 "Example Video") from Worksheet (illustrating code in GetStoriesForm).

## Code Structure

The core functionality is contained in the Class Modules that are accessible from the VBA Editor for the Excel Worksheet:

![Core Classes](https://raw.githubusercontent.com/markwilliams970/RallyRestToolkitForVBA/master/screenshots/screenshot7.png)

## Example Code

The code behind the "Get Stories" button in the worksheet is contained within the "GetStoriesForm" module. There's some UI logic there, but the useful stuff around how to instantiate and use the RallyRestToolkitForVBA toolkit is in the QueryStories function:

![Sample Code](https://raw.githubusercontent.com/markwilliams970/RallyRestToolkitForVBA/master/screenshots/screenshot8.png)

There is an example showing how to use the RallyRestToolkitForVBA toolkit to Create Defects. The data and upload button are on the "CreateDefects" Worksheet.

![CreateDefects](https://raw.githubusercontent.com/markwilliams970/RallyRestToolkitForVBA/master/screenshots/screenshot10.png)

 The code that accomplishes the upload is found within the "UploadDefectsForm" module.

![UploadDefectsForm](https://raw.githubusercontent.com/markwilliams970/RallyRestToolkitForVBA/master/screenshots/screenshot11.png)

There's an example data and button for Updating Defects present on the "UpdateDefects" Worksheet Tab:

![UpdateDefects](https://raw.githubusercontent.com/markwilliams970/RallyRestToolkitForVBA/master/screenshots/screenshot12.png)

The code that accomplishes the updating is found within the "UpdateDefectsForm" Module:

![UpdateDefectsForm](https://raw.githubusercontent.com/markwilliams970/RallyRestToolkitForVBA/master/screenshots/screenshot13.png)

## To Do

There's lots still to do with this to make this toolkit robust. Overall, error-checking and handling needs to be a lot more thorough everywhere. There are probably a lot of situations where text that resides within Excel cells will require more complete encoding and escaping before uploading to Rally, in order to get things to work right. This is alpha-level code...so just be aware.

## Will this work with Excel for Macs?

Unfortunately, No. The MSXML2 module that VBA uses for the HTTP connection is not available on the Mac.
