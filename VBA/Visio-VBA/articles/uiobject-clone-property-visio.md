---
title: UIObject.Clone Property (Visio)
keywords: vis_sdr.chm14913245
f1_keywords:
- vis_sdr.chm14913245
ms.prod: visio
api_name:
- Visio.UIObject.Clone
ms.assetid: 9fd3eb9b-8b01-9397-8f9f-58e3ce4a980e
ms.date: 06/08/2017
---


# UIObject.Clone Property (Visio)

Returns a copy of the  **UIObject** object. Read-only.


## Syntax

 _expression_ . **Clone**

 _expression_ A variable that represents a **UIObject** object.


### Return Value

UIObject


## Example

This example shows how to use the  **Clone** property to make a copy of a **UIObject** object. The example adds a custom toolbar to the cloned toolbars collection. This toolbar appears in the Microsoft Visio user interface and is available while the document is active.

To restore the built-in toolbars in Visio after you run this macro, call the  **ThisDocument.ClearCustomToolbars** method.






```vb
 
Sub Clone_Example() 
 
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
 Set vsoToolbars = vsoUIObject.ToolbarSets.ItemAtID(Visio.visUIObjSetDrawing).Toolbars 
 
 'Add a toolbar to the collection. 
 Set vsoToolbar = vsoToolbars.Add 
 
 'Set the title of the toolbar. 
 vsoToolbar.Caption = "My New Toolbar" 
 
 'Tell Visio to use the new UIObject object while 
 'this document is active. 
 ThisDocument.SetCustomToolbars vsoUIObject 
 
End Sub
```


