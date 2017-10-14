---
title: ToolbarItem.Style Property (Visio)
keywords: vis_sdr.chm13551150
f1_keywords:
- vis_sdr.chm13551150
ms.prod: visio
api_name:
- Visio.ToolbarItem.Style
ms.assetid: 3f4c624c-a80c-9b82-3ca5-34c395770bf0
ms.date: 06/08/2017
---


# ToolbarItem.Style Property (Visio)

Determines whether a toolbar button shows an icon, a caption, or some combination. Read/write.


## Syntax

 _expression_ . **Style**

 _expression_ A variable that represents a **ToolbarItem** object.


### Return Value

Integer


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

Possible values for the  **Style** property are listed in the following table. These constants are declared by the Visio type library in **VisUIButtonStyle** .



|** Constant**|** Value**|
|:-----|:-----|
| **visButtonAutomatic**|0|
| **visButtonCaption**|1|
| **visButtonIcon**|2|
| **visButtonIconandCaption**|3|

## Example

This example shows how to use the  **Style** property to set the style of a toolbar button. The example adds a custom toolbar button and sets it to display both an icon and a caption. This button appears in the Visio user interface and is available while the document is active.

Before running this code, replace  _path\filename_ with the full path to and name of a valid icone (.ico) file on your computer.

To restore the built-in toolbars in Microsoft Visio after you run this macro, call the  **ThisDocument.ClearCustomToolbars** method.




```vb
Sub Style_Example() 
 
 Dim vsoUIObject As Visio.UIObject 
 Dim vsoToolbarSet As Visio.ToolbarSet 
 Dim vsoToolbarItems As Visio.ToolbarItems 
 Dim vsoToolbarItem As Visio.ToolbarItem 
 
 'Check whether there are document custom toolbars. 
 If ThisDocument.CustomToolbars Is Nothing Then 
 
 'If not, check whether there are application custom toolbars. 
 If Visio.Application.CustomToolbars Is Nothing Then 
 
 'If not, use the built-in toolbars. 
 Set vsoUIObject = Visio.Application.BuiltInToolbars(0) 
 
 Else 
 
 'If there are application custom toolbars, copy them. 
 Set vsoUIObject = Visio.Application.CustomToolbars.Clone 
 End If 
 
 Else 
 
 'If there already are document custom toolbars, use them. 
 Set vsoUIObject = ThisDocument.CustomToolbars 
 
 End If 
 
 'Get the Toolbars collection for the drawing window context. 
 Set vsoToolbarSet = vsoUIObject.ToolbarSets.ItemAtID(visUIObjSetDrawing) 
 
 'Get the set of toolbar items for the Standard toolbar. 
 Set vsoToolbarItems = vsoToolbarSet.Toolbars(0).ToolbarItems 
 
 'Add a new button in the first position. 
 Set vsoToolbarItem = vsoToolbarItems.AddAt(0) 
 
 'Set properties for the new toolbar button. 
 vsoToolbarItem.CntrlType = visCtrlTypeBUTTON 
 vsoToolbarItem.CmdNum = visCmdFileSave 
 vsoToolbarItem.Style = visButtonIconandCaption 
 vsoToolbarItem.IconFileName "path\filename " 
 
 'Use the new UIObject object while this document is active. 
 ThisDocument.SetCustomToolbars vsoUIObject 
 
End Sub
```


