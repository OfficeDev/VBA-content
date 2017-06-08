---
title: ToolbarItem.ToolbarItems Property (Visio)
keywords: vis_sdr.chm13514550
f1_keywords:
- vis_sdr.chm13514550
ms.prod: visio
api_name:
- Visio.ToolbarItem.ToolbarItems
ms.assetid: 5bc9b339-ffc8-1aef-526e-a815c3c5e251
ms.date: 06/08/2017
---


# ToolbarItem.ToolbarItems Property (Visio)

Returns the  **ToolbarItems** collection of a **ToolbarItem** object. Read-only.


## Syntax

 _expression_ . **ToolbarItems**

 _expression_ A variable that represents a **ToolbarItem** object.


### Return Value

ToolbarItems


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **ToolbarItems** property to get a particular object in a collection. It also shows how to get a copy of the built-in Visio toolbars, add a toolbar button, set the button icon, and replace the built-in toolbar set with the custom set.

Before running this code, replace  _path\filename_ with the full path to and name of a valid icon (.ico) file on your computer.

To restore the built-in Visio toolbars after you run this macro, call the  **ThisDocument.ClearCustomToolbars** method.




```vb
 
Public Sub ToolbarItems_Example() 
 
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


