---
title: ToolbarItem.State Property (Visio)
keywords: vis_sdr.chm13514425
f1_keywords:
- vis_sdr.chm13514425
ms.prod: visio
api_name:
- Visio.ToolbarItem.State
ms.assetid: 1a3c97f0-d6bd-00d2-0caf-ac9876c533aa
ms.date: 06/08/2017
---


# ToolbarItem.State Property (Visio)

Determines a button's state, pressed or not pressed. Read/write.


## Syntax

 _expression_ . **State**

 _expression_ A variable that represents a **ToolbarItem** object.


### Return Value

Integer


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

The  **State** property can be one of the following constants declared by the Visio type library in **VisUIButtonState** .



|** Constant**|** Value**|** Description**|
|:-----|:-----|:-----|
| **visButtonUp**|0|Button is not pressed|
| **visButtonDown**|-1|Button is pressed|

## Example

This example shows how to use the  **State** property to set the state of a toolbar button (make it appear pressed). The example adds a custom toolbar button to the **Standard** toolbar. Pressing the button saves the active document. This button appears in the Microsoft Visio user interface and is available while the document is active.

Before running this code, replace  _path\filename_ with the full path to and name of a valid icon (.ico) file on your computer.

To restore the built-in Visio toolbars after you run this macro, call the  **ThisDocument.ClearCustomToolbars** method.




```vb
 
Sub State_Example() 
 
 Dim vsoUIObject As Visio.UIObject 
 Dim vsoToolbarSet As Visio.ToolbarSet 
 Dim vsoToolbarItems As Visio.ToolbarItems 
 Dim vsoToolbarItem As Visio.ToolbarItem 
 
 'Check whether there are document custom toolbars. 
 If ThisDocument.CustomToolbars Is Nothing Then 
 
 'If not, check whether there are application custom toolbars. 
 If Visio.Application.CustomToolbars Is Nothing Then 
 
 'If there are no custom toolbars, use the built-in toolbars. 
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
 vsoToolbarItem.State = visButtonDown 
 vsoToolbarItem.IconFileName "path\filename " 
 
 
 'Use the new UIObject object while this document is active. 
 ThisDocument.SetCustomToolbars vsoUIObject 
 
End Sub
```


