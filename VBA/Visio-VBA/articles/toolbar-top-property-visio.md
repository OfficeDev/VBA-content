---
title: Toolbar.Top Property (Visio)
keywords: vis_sdr.chm13714570
f1_keywords:
- vis_sdr.chm13714570
ms.prod: visio
api_name:
- Visio.Toolbar.Top
ms.assetid: 63adeae5-c962-4e83-67de-d89035ee9bce
ms.date: 06/08/2017
---


# Toolbar.Top Property (Visio)

Gets the distance between the top of an object and the top of the docking area or the top of the screen if the object isn't docked; it sets the distance between the top of a  **Toolbar** object and the top of the screen. Read/write.


## Syntax

 _expression_ . **Top**

 _expression_ A variable that represents a **Toolbar** object.


### Return Value

Integer


## Example

This example shows how to use the  **Top** property to set the position of a **UIObject** object. The example adds a custom toolbar to the cloned toolbars collection. This toolbar appears in the Microsoft Visio user interface and is available while the document is active.

To restore Visio's built-in toolbars after you run this macro, call the  **ThisDocument.ClearCustomToolbars** method.




```vb
 
Public Sub Top_Example() 
 
 Dim vsoUIObject As Visio.UIObject 
 Dim vsoToolbars As Visio.Toolbars 
 Dim vsoToolbar As Visio.Toolbar 
 
 'Check whether there are document custom toolbars. 
 If ThisDocument.CustomToolbars Is Nothing Then 
 
 'If not, check whether there are application custom toolbars. 
 If Visio.Application.CustomToolbars Is Nothing Then 
 
 'If not, use the built-in toolbars. 
 Set vsoUIObject = Visio.Application.BuiltInToolbars(0) 
 
 Else 
 
 'If there are application custom toolbars, clone them. 
 Set vsoUIObject = Visio.Application.CustomToolbars.Clone 
 
 End If 
 
 Else 
 
 'If there are custom toolbars in the document, use them. 
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
 
 End With 
 
 'Use the custom toolbars in this document. 
 ThisDocument.SetCustomToolbars vsoUIObject 
 
End Sub
```


