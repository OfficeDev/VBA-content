---
title: Set Visual Basic Environment Options
keywords: vbhw6.chm1105240
f1_keywords:
- vbhw6.chm1105240
ms.prod: office
ms.assetid: ce85ae8c-9e02-2525-98e7-403d5a590d6c
ms.date: 06/08/2017
---


# Set Visual Basic Environment Options

You can set the behavior and look of the Visual Basic development environment through the  **Options** dialog box. Use the:



-  **Editor** tab to specify Code window and Project window settings.

-  **Editor Format** tab to specify the appearance of your code.

-  **Genera** l tab to specify form, error handling, and compile settings for your project.

-  **Docking** tab to specify whether a window is attached or "anchored" to one edge of other dockable or application windows.


 **To set Environment options**


- On the  **Tools** menu of the Visual Basic editor, click **Options**. Each option is described in the following tables.


 **Editor**


| <strong>Option</strong>                       | <strong>Description</strong>                                                                                                                             |
|:----------------------------------------------|:---------------------------------------------------------------------------------------------------------------------------------------------------------|
| <strong>Auto Syntax Check</strong>            | Visual Basic automatically verifies correct syntax after you enter a line of code.                                                                       |
| <strong>Require Variable Declaration</strong> | Explicit variable declarations are required in [modules](vbe-glossary.md).                                                                               |
| <strong>Auto Indent</strong>                  | After tabbing the first line of code, all subsequent lines start at that tab location.                                                                   |
| <strong>Tab Width</strong>                    | The tab width, which can range from 1 - 32 spaces. (Default is 4 spaces.)                                                                                |
| <strong>Default to Full Module View</strong>  | [Procedures](vbe-glossary.md) for new modules are displayed in the <strong>Code</strong> window as a single, scrollable list or one procedure at a time. |
| <strong>Procedure Separator</strong>          | Display separator bars at the end of each procedure in the  <strong>Code</strong> window.                                                                |
| <strong>Auto List Members</strong>            | At the insertion point, Visual Basic displays information that logically completes a statement.                                                          |
| <strong>Auto Quick Info</strong>              | Information about functions and their [arguments](vbe-glossary.md) is displayed as you type.                                                             |
| <strong>Auto Data Tips</strong>               | Automatically display the value of any [variable](vbe-glossary.md) on which you place the mouse pointer. Available only in[break mode](vbe-glossary.md). |
| <strong>Drag-Drop in Text Editing</strong>    | Code elements can be dragged from the  <strong>Code</strong> window into the <strong>Immediate</strong> or <strong>Watch</strong> windows.               |

 **Editor Format**


| <strong>Option</strong>                                                                  | <strong>Description</strong>                                                                |
|:-----------------------------------------------------------------------------------------|:--------------------------------------------------------------------------------------------|
| <strong>Foreground</strong>, <strong>Background</strong>, and <strong>Indicator</strong> | The color of different categories of text listed in the  <strong>Code Colors</strong> list. |
| <strong>Font</strong>                                                                    | The font used for displaying code.                                                          |
| <strong>Size</strong>                                                                    | The size of the font used for code.                                                         |
| <strong>Margin Indicator Bar</strong>                                                    | Display the  <strong>Margin Indicator Bar</strong>.                                         |

 **General**


| <strong>Option</strong>                       | <strong>Description</strong>                                                                                                                                                   |
|:----------------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <strong>Show Grid</strong>                    | Display a grid on a form.                                                                                                                                                      |
| <strong>Grid Units</strong>                   | Lists the unit of measurement for units in the grid.                                                                                                                           |
| <strong>Width</strong>                        | The width of the grid cells on a form.                                                                                                                                         |
| <strong>Height</strong>                       | The height of the grid cells on a form.                                                                                                                                        |
| <strong>Align Controls to Grid</strong>       | Automatically position the outer edge of controls on the closest grid lines.                                                                                                   |
| <strong>Show ToolTips</strong>                | Display ToolTips for toolbar buttons.                                                                                                                                          |
| <strong>Collapse Proj. Hides Windows</strong> | Automatically close the project,  <strong>UserForm</strong>, object, or module windows when a[project](vbe-glossary.md) is collapsed in the <strong>Project Explorer</strong>. |
| <strong>Notify Before State Loss</strong>     | Display a message that a requested action will cause all module-level variables to be reset for a running project.                                                             |
| <strong>Break on All Errors</strong>          | Any error causes the project to enter break mode, whether or not an error handler is active, and whether or not the code is in a [class module](vbe-glossary.md).              |
| <strong>Break in Class Module</strong>        | Any unhandled error produced in a class module causes the project to enter break mode at the line of code which produced the error.                                            |
| <strong>Break on Unhandled Errors</strong>    | Any other unhandled error causes the project to enter break mode.                                                                                                              |
| <strong>Compile On Demand</strong>            | A project is fully compiled before it starts, or code is compiled as needed.                                                                                                   |
| <strong>Background Compile</strong>           | Use idle time during run time to finish compiling the project in the background. (Available only if  <strong>Compile On Demand</strong> is set.)                               |

 **Docking**


|**Option**|**Description**|
|:-----|:-----|
|The check box for the appropriate window|A window can be anchored to an adjacent dockable window or the Visual Basic Editor window.|

