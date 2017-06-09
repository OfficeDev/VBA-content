---
title: Toolbar.Left Property (Visio)
keywords: vis_sdr.chm13713825
f1_keywords:
- vis_sdr.chm13713825
ms.prod: visio
api_name:
- Visio.Toolbar.Left
ms.assetid: 2929fef2-0855-dae1-9c60-0208d1de4dee
ms.date: 06/08/2017
---


# Toolbar.Left Property (Visio)

Gets the distance in pixels between the left edge of the object and the left side of the docking area. Sets the distance in pixels between the left edge of an object and the left edge of the screen. Read/write.


## Syntax

 _expression_ . **Left**

 _expression_ A variable that represents a **Toolbar** object.


### Return Value

Integer


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

The value of  **Left** must be greater than or equal to zero.


## Example

This example shows how to use the  **Left** property to set the position of a **UIObject** object. The example adds a custom toolbar to the cloned toolbars collection. This toolbar appears in the Microsoft Visio user interface and is available while the document is active.

To restore Visio's built-in toolbars after you run this macro, call the  **ThisDocument.ClearCustomToolbars** method.




```vb
 
Sub Left_Example() 
 
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
 Set VsoToolbars = vsoUIObject.ToolbarSets.ItemAtID( _ 
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
 
 End With 
 
 'Tell Microsoft Office Visio to use the new UIObject object while 
 'this document is active. 
 ThisDocument.SetCustomToolbars vsoUIObject 
 
End Sub
```


