---
title: Application.BuiltInToolbars Property (Visio)
keywords: vis_sdr.chm10013165
f1_keywords:
- vis_sdr.chm10013165
ms.prod: visio
api_name:
- Visio.Application.BuiltInToolbars
ms.assetid: e0460fa5-23da-f452-f541-feabe8e3bffb
ms.date: 06/08/2017
---


# Application.BuiltInToolbars Property (Visio)

Returns a  **UIObject** object that represents a copy of the built-in Microsoft Visio toolbars. Read-only.


## Syntax

 _expression_ . **BuiltInToolbars**( **_fIgnored_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _fIgnored_|Required| **Integer**| Required **Integer** . Required for backward compatibility, but ignored. It is suggested that you pass zero (0).|

### Return Value

UIObject


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

You can use the  **BuiltInToolbars** property to obtain a **UIObject** object and modify its toolbars. You can then use the **SetCustomToolbars** method of an **Application** or **Document** object to substitute your customized toolbars for the built-in Visio toolbars.

You can also use the  **SaveToFile** method of the **UIObject** object to store its toolbars in a file and reload them as custom toolbars by setting the **CustomToolbarsFile** property of an **Application** or **Document** object.

Prior to Visio 5.0, the argument for this property was ( _fWhichToolbars_), which designated the type of toolbar to get (MSOffice or LotusSS). Beginning with Visio 5.0, the application no longer supports different types of toolbars and the current argument, ( _fIgnored_), is ignored.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **BuiltInToolbars** property to get a copy of the built-in Visio toolbars, add a toolbar button, set the button icon, and replace the built-in toolbar set with the custom set.

Before running this macro, replace  _(path\filename)_ in the code below with the full path to and file name of an icon file (.ico) on your computer.

To restore the built in Visio toolbars after you run this macro, call the  **ThisDocument.ClearCustomToolbars** method.




```vb
 
Public Sub BuiltInToolbars_Example() 
 
 Dim vsoUIObject As Visio.UIObject 
 Dim vsoToolbarSet As Visio.ToolbarSet 
 Dim vsoToolbarItems As Visio.ToolbarItems 
 Dim vsoToolbarItem As Visio.ToolbarItem 
 
 'Get the UIObject object for the copy of the built-in toolbars. 
 Set vsoUIObject = Visio.Application.BuiltInToolbars(0) 
 
 'Get the drawing window toolbar sets.'NOTE: Use ItemAtID to get the toolbar sets.'Using vsoUIObject.ToolbarSets(visUIObjSetDrawing) will not work. 
 Set vsoToolbarSet = vsoUIObject.ToolbarSets.ItemAtID(visUIObjSetDrawing) 
 
 'Get the ToolbarItems collection. 
 Set vsoToolbarItems = vsoToolbarSet.Toolbars(0).ToolbarItems 
 
 'Add a new button in the first position. 
 Set vsoToolbarItem = vsoToolbarItems.AddAt(0) 
 
 'Set properties for the new toolbar button. 
 vsoToolbarItem.CntrlType = visCtrlTypeBUTTON 
 
 'Set the caption for the new toolbar button. 
 vsoToolbarItem.Caption = "MyButton" 
 
 'Set the icon for the new toolbar button. 
 vsoToolbarItem.IconFileName "(path\filename)" 'Tell Visio to actually use the new custom UI. 
 ThisDocument.SetCustomToolbars vsoUIObject 
 
End Sub
```


