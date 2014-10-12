RallyRestToolkitForVBA
======================

##What Is It?

A Rally REST toolkit written for Microsoft Visual Basic for Applications (VBA). VBA is a scripting language used by Microsoft Office (and other) applications.

##Why?
Why not? VBA is a useful toolkit for automating processes in Microsoft Excel, for one. The [Rally Excel Plugin](https://help.rallydev.com/rally-add-excel "Rally Excel Add-In"), while providing a nice UI-based interface for querying Rally, does not offer script-ability or automation.

##Why Not?
Excellent question. VBA may be quite slow and inefficient for accessing large amounts of data. It's not asynchronous, and can be a clumsy way to get data into/out of Rally in large volumes. But for small datasets, it could be convenient.

##How functional is this toolkit at this point in time?
Barely. Pre-alpha. Right now this is nothing more than a proof-of-concept for Querying Rally and Creating Rally Artifacts from VBA code. More to come though!

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

## Code Structure

The core functionality is contained in the Class Modules that are accessible from the VBA Editor for the Excel Worksheet:

![Core Classes](https://raw.githubusercontent.com/markwilliams970/RallyRestToolkitForVBA/master/screenshots/screenshot7.png)

## Example Code

The code behind the "Get Stories" button in the worksheet is contained within the "GetStoriesForm" module. There's some UI logic there, but the useful stuff around how to instantiate and use the RallyRestToolkitForVBA toolkit is in the QueryStories function:

![Sample Code](https://raw.githubusercontent.com/markwilliams970/RallyRestToolkitForVBA/master/screenshots/screenshot8.png)

There is also an example showing how to use the RallyRestToolkitForVBA toolkit to Create Defects, within the "Examples" module. This sample isn't hooked up to any UI Components.

![Create Defects Example](https://raw.githubusercontent.com/markwilliams970/RallyRestToolkitForVBA/master/screenshots/screenshot9.png)

## To Do

There's lots still to do with this to make this toolkit anything close to useful. There's the U/D of CRUD (Update and Delete) still to do. Overall, error-checking and handling needs to be a lot more robust throughout. There are probably a lot of situations where text that resides within Excel cells will require more thorough encoding and escaping before uploading to Rally, in order to get things to work right. This is alpha-level code...so just be aware.

## Will this work with Excel for Macs?

Unfortunately, No. The MSXML2 module that VBA uses for the HTTP connection is not available on the Mac.
