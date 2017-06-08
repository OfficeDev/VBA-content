---
title: Toolbar.Position Property (Visio)
keywords: vis_sdr.chm13714095
f1_keywords:
- vis_sdr.chm13714095
ms.prod: visio
api_name:
- Visio.Toolbar.Position
ms.assetid: a1642793-7e72-332e-db3c-67438ac62675
ms.date: 06/08/2017
---


# Toolbar.Position Property (Visio)

Gets or sets the position of an object. Read/write.


## Syntax

 _expression_ . **Position**

 _expression_ A variable that represents a **Toolbar** object.


### Return Value

Integer


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

Constants that represent possible  **Position** property values are listed below. They are also declared by the Visio type library in **VisUIBarPosition** .



|** Constant**|** Value**|
|:-----|:-----|
| **visBarLeft**|0|
| **visBarTop**|1|
| **visBarRight**|2|
| **visBarBottom**|3|
| **visBarFloating**|4|
| **visBarPopup**|5|
| **visBarMenu**|6|

## Example

This example shows how to use the  **Position** property to set the position of a custom toolbar. The example adds a custom toolbar to the **Toolbars** collection. The toolbar appears in the Visio user interface and is available while the document is active.

To restore the built-in toolbars in Microsoft Visio after you run this macro, call the  **ThisDocument.ClearCustomToolbars** method.




```vb
 
Sub Position_Example() 
 
 Dim vsoUIObject As Visio.UIObject 
 Dim vsoToolbars As Visio.Toolbars 
 Dim vsoToolbar As Visio.Toolbar 
 
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
 .Caption = "Test" 
 
 'Float the toolbar at coordinates (300,200). 
 .Position = Visio.visBarFloating 
 .Left = 300 
 .Top = 200 
 
 'Disallow docking the new toolbar. 
 .Protection = Visio.visBarNoHorizontalDock _ 
 + Visio.visBarNoVerticalDock 
 
 End With 
 
 'Use the new UIObject object while 
 'this document is active. 
 ThisDocument.SetCustomToolbars vsoUIObject 
 
End Sub
```


