---
title: UIObject.ToolbarSets Property (Visio)
keywords: vis_sdr.chm14914560
f1_keywords:
- vis_sdr.chm14914560
ms.prod: visio
api_name:
- Visio.UIObject.ToolbarSets
ms.assetid: 5fd4551c-3e23-920b-9dbc-76b2a79671f4
ms.date: 06/08/2017
---


# UIObject.ToolbarSets Property (Visio)

Returns the  **ToolbarSets** collection of a **UIObject** object. Read-only.


## Syntax

 _expression_ . **ToolbarSets**

 _expression_ A variable that represents a **UIObject** object.


### Return Value

ToolbarSets


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

If a  **UIObject** object represents toolbars (for example, if the object was retrieved by using the **BuiltInToolbars** property of an **Application** object), its **ToolbarSets** collection represents all of the toolbars for that **UIObject** object.

Use the  **ItemAtID** property of a **ToolbarSets** object to retrieve toolbars for a particular window context, for example, the drawing window. If a context does not include toolbars, it has no **ToolbarSets** collection.


## Example

This Microsoft Visual Basic macro shows how to use the  **ToolbarSets** property to get a particular object in a collection. It also shows how to get a copy of the built-in Visio toolbars, add a toolbar button, set the button icon, and replace the built-in toolbar set with the custom set.



Before running this code, replace  _path\filename_ with the full path to and name of a valid icon (.ico) file on your computer.

To restore the built-in Visio toolbars after you run this macro, call the  **ThisDocument.ClearCustomToolbars** method.




```vb
 
Public Sub ToolbarSets_Example() 
 
 Dim vsoUIObject As Visio.UIObject 
 Dim vsoToolbarSet As Visio.ToolbarSet 
 Dim vsoToolbarItems As Visio.ToolbarItems 
 Dim vsoToolbarItem As Visio.ToolbarItem 
 
 'Get the UIObject object for the copy of the built-in toolbars. 
 Set vsoUIObject = Visio.Application.BuiltInToolbars(0) 
 
 'Get the drawing window toolbar sets. 
 'NOTE: Use ItemAtID to get the toolbar set. 
 'Using vsoUIObject.ToolbarSets(visUIObjSetDrawing) will not work. 
 Set vsoToolbarSet = vsoUIObject.ToolbarSets.ItemAtID(visUIObjSetDrawing) 
 
 'Get the ToolbarItems collection. 
 Set vsoToolbarItems = vsoToolbarSet.Toolbars(0).ToolbarItems 
 
 'Add a new button in the first position. 
 Set vsoToolbarItem = vsoToolbarItems.AddAt(0) 
 
 'Set properties for the new toolbar button. 
 vsoToolbarItem.CntrlType = visCtrlTypeBUTTON 
 vsoToolbarItem.CmdNum = visCmdPanZoom 
 
 'Set the toolbar button icon. 
 vsoToolbarItem.IconFileName "path\filename " 
 
 'Use the new custom UI. 
 ThisDocument.SetCustomToolbars vsoUIObject 
 
End Sub
```


