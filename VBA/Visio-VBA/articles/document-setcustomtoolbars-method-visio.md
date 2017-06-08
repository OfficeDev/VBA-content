---
title: Document.SetCustomToolbars Method (Visio)
keywords: vis_sdr.chm10516565
f1_keywords:
- vis_sdr.chm10516565
ms.prod: visio
api_name:
- Visio.Document.SetCustomToolbars
ms.assetid: fddae53c-0519-90ef-0023-ee3896e86757
ms.date: 06/08/2017
---


# Document.SetCustomToolbars Method (Visio)

Replaces the current built-in or custom toolbars of an application or document.


## Syntax

 _expression_ . **SetCustomToolbars**( **_ToolbarsObject_** )

 _expression_ A variable that represents a **Document** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ToolbarsObject_|Required| **[IVUIOBJECT]**|An expression that returns a  **UIObject** object that represents the new custom toolbars.|

### Return Value

Nothing


## Remarks

If the  _ToolbarsObject_ object was created in a separate process by using the VBA **CreateObject** method instead of getting the appropriate property of an **Application** or **Document** object, the **SetCustomToolbars** method returns an error.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **SetCustomToolbars** method to replace the built-in toolbar set with the custom set. It retrieves a copy of the built-in Visio toolbars, adds a toolbar and a toolbar button, sets the button icon, and then replaces the toolbar set.

Before running this macro, replace  _path\filename_ in the code with the full path to and filename of an icon file (.ico) on your computer.




```vb
Public Sub SetCustomToolbarItems_Example() 
 
 Dim vsoUIObject As Visio.UIObject 
 Dim vsoToolbarSet As Visio.ToolbarSet 
 Dim vsoToolbar As Visio.Toolbar 
 Dim vsoToolbarItems As Visio.ToolbarItems 
 Dim vsoToolbarItem As Visio.ToolbarItem 
 
 'Get the UIObject object for the copy of the built-in toolbars. 
 Set vsoUIObject = Visio.Application.BuiltInToolbars(0) 
 
 'Get the drawing window toolbar sets. 
 'NOTE: Use ItemAtID to get the toolbar set. 
 'Using vsoUIObject.ToolbarSets(visUIObjSetDrawing) will not work. 
 Set vsoToolbarSet = vsoUIObject.ToolbarSets.ItemAtID(visUIObjSetDrawing) 
 
 'Create a new toolbar 
 Set vsoToolbar = vsoToolbarSet.Toolbars.Add 
 
 With vsoToolbar 
 .Caption = "test" 
 .Position = visBarFloating 
 .Left = 300 
 .Top = 200 
 
 .Protection = visBarNoHorizontalDock 
 .Visible = True 
 .Enabled = True 
 End With 
 
 'Get the ToolbarItems collection. 
 Set vsoToolbarItems = vsoToolbar.ToolbarItems 
 
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


