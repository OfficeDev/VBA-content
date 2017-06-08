---
title: ToolbarSets.ItemAtID Property (Visio)
keywords: vis_sdr.chm14013770
f1_keywords:
- vis_sdr.chm14013770
ms.prod: visio
api_name:
- Visio.ToolbarSets.ItemAtID
ms.assetid: 5508ee05-03ca-547d-26dc-2b80c0c22f49
ms.date: 06/08/2017
---


# ToolbarSets.ItemAtID Property (Visio)

Returns the  **ToolbarSet** object for an ID within a collection. Read-only.


## Syntax

 _expression_ . **ItemAtID**( **_lID_** )

 _expression_ A variable that represents a **ToolbarSets** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _lID_|Required| **Long**|The Visio context ID of the object to retrieve.|

### Return Value

ToolbarSet


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

The ID corresponds to a window or context menu. Constants for IDs are prefixed with  **visUIObjectSet** and are declared by the Visio type library in **[VisUIObjSets](visuiobjsets-enumeration-visio.md)** .


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **ItemAtID** property to get a particular object in a collection. It also shows how to get a copy of the built-in Visio toolbars, add a toolbar button, set the button icon, and replace the built-in toolbar set with the custom set.

Before running this code, replace  _path_ \ _filename_ with the full path to and name of a valid icon (.ico) file on your computer.

To restore the built-in Visio toolbars after you run this macro, call the  **ThisDocument.ClearCustomToolbars** method.




```vb
Public Sub ItemAtID_Example() 
 
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
 
 'Set the toolbar button caption. 
 vsoToolbarItem.Caption = "MyButton" 
 
 'Set the toolbar button icon. 
 vsoToolbarItem.IconFileName "path \filename " 
 
 'Tell Visio to actually use the new custom UI. 
 ThisDocument.SetCustomToolbars vsoUIObject 
 
End Sub
```


