
# MenuSet.Protection Property (Visio)

 **Last modified:** July 28, 2015

 _**Applies to:** Visio 2013 Preview_

Determines how a  **MenuSet** object is protected from user customization. Read/write.


## Syntax

 _expression_. **Protection**

 _expression_A variable that represents a  **MenuSet** object.


### Return Value

Integer


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

The value of the  **Protection** property can be one or a combination of the following constants declared by the Visio type library in **VisUIBarProtection**.



|** Constant**|** Value**|** Description**|
|:-----|:-----|:-----|
| **visBarNoProtection**|0|No protection.|
| **visBarNoCustomize**|1|Cannot be customized.|
| **visBarNoResize**|2|Cannot be resized.|
| **visBarNoMove**|4|Cannot be moved.|
| **visBarNoChangeDock**|16|Cannot be docked or floating.|
| **visBarNoVerticalDock**|32|Cannot be docked vertically.|
| **visBarNoHorizontalDock**|64|Cannot be docked horizontally.|
