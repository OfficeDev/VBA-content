---
title: ToolbarItem.Visible Property (Visio)
keywords: vis_sdr.chm13514650
f1_keywords:
- vis_sdr.chm13514650
ms.prod: visio
api_name:
- Visio.ToolbarItem.Visible
ms.assetid: 1fe7078b-1e8a-da95-7289-d1d83f441f67
ms.date: 06/08/2017
---


# ToolbarItem.Visible Property (Visio)

Determines whether an object is visible. Read/write.


## Syntax

 _expression_ . **Visible**

 _expression_ A variable that represents a **ToolbarItem** object.


### Return Value

Boolean


## Example

This example shows how to use the  **Visible** property to determine if a **UIObject** object is visible in the user interface. The example adds a custom toolbar to the cloned toolbars collection. This toolbar appears in the Microsoft Visio user interface and is available while the document is active.

To restore the built-in Visio toolbars after you run this macro, call the  **ThisDocument.ClearCustomToolbars** method.




```vb
 
Public Sub Visible_Example() 
 
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
 
 'Make the toolbar visible. 
 .Visible = True 
 
 End With 
 
 'Use the custom toolbars in this document. 
 ThisDocument.SetCustomToolbars vsoUIObject 
 
End Sub
```


