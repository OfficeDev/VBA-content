---
title: Toolbar.Protection Property (Visio)
keywords: vis_sdr.chm13714180
f1_keywords:
- vis_sdr.chm13714180
ms.prod: visio
api_name:
- Visio.Toolbar.Protection
ms.assetid: 2f2120db-78de-d37c-4764-c3fabe17a6f5
ms.date: 06/08/2017
---


# Toolbar.Protection Property (Visio)

Determines how a  **Toolbar** object is protected from user customization. Read/write.


## Syntax

 _expression_ . **Protection**

 _expression_ A variable that represents a **Toolbar** object.


### Return Value

Integer


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

The value of the  **Protection** property can be one or a combination of the following constants declared by the Visio type library in **VisUIBarProtection** .



|** Constant**|** Value**|** Description**|
|:-----|:-----|:-----|
| **visBarNoProtection**|0|No protection.|
| **visBarNoCustomize**|1|Cannot be customized.|
| **visBarNoResize**|2|Cannot be resized.|
| **visBarNoMove**|4|Cannot be moved.|
| **visBarNoChangeDock**|16|Cannot be docked or floating.|
| **visBarNoVerticalDock**|32|Cannot be docked vertically.|
| **visBarNoHorizontalDock**|64|Cannot be docked horizontally.|

## Example

This example shows how to use the  **Protection** property to prevent users from docking a custom toolbar. The example adds a custom toolbar to the **Toolbars** collection and then adds a button to the toolbar. The toolbar appears in the Microsoft Visio user interface and is available while the document is active.

To restore Visio's built-in toolbars after you run this macro, call the  **ThisDocument.ClearCustomToolbars** method.




```vb
 
Sub Protection_Example() 
 
 Dim vsoUIObject As Visio.UIObject 
 Dim vsoToolbars As Visio.Toolbars 
 Dim vsoToolbar As Visio.Toolbar 
 Dim vsoToolbarItem As Visio.ToolbarItem 
 
 'Check whether there are document custom toolbars. 
 If ThisDocument.CustomToolbars Is Nothing Then 
 
 'Check whether there are application custom toolbars. 
 If Visio.Application.CustomToolbars Is Nothing Then 
 
 'Use the built-in toolbars. 
 Set vsoUIObject = Visio.Application.BuiltInToolbars(0) 
 
 Else 
 
 'Use the application custom toolbars. 
 Set vsoUIObject = Visio.Application.CustomToolbars.Clone 
 
 End If 
 
 Else 
 
 'Use the document custom toolbars. 
 Set vsoUIObject = ThisDocument.CustomToolbars 
 
 End If 
 
 'Get the Toolbars collection for the drawing window context. 
 Set vsoToolbars = vsoUIObject.ToolbarSets.ItemAtID( _ 
 Visio.visUIObjSetDrawing).Toolbars 
 
 'Add a toolbar to the collection. 
 Set vsoToolbar = vsoToolbars.Add 
 With vsoToolbar 
 
 'Set the title of the toolbar. 
 .Caption = "My New Toolbar" 
 
 'Float the toolbar at coordinates (300,200). 
 .Position = Visio.visBarFloating 
 .Left = 300 
 .Top = 200 
 
 'Disallow docking the new toolbar. 
 .Protection = Visio.visBarNoHorizontalDock _ 
 + Visio.visBarNoVerticalDock 
 
 End With 
 
 'Add an item to the toolbar. 
 Set vsoToolbarItem = vsoToolbar.ToolbarItems.Add 
 With vsoToolbarItem 
 
 'Set the new item to be a button. 
 .CntrlType = Visio.visCtrlTypeBUTTON 
 
 'Set the icon of the new button. 
 .FaceID = Visio.visIconIXCUSTOM_CARDS 
 
 'Set the CmdNum property of the new button 
 .CmdNum = 1 
 
 'Set the Width property of the new button 
 'wide enough that the toolbar name is readable. 
 .Width = 100 
 
 End With 
 
 'Tell Visio to use the new UIObject object while 
 'this document is active. 
 ThisDocument.SetCustomToolbars vsoUIObject 
 
End Sub
```


