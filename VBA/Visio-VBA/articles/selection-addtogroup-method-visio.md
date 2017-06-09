---
title: Selection.AddToGroup Method (Visio)
keywords: vis_sdr.chm11116070
f1_keywords:
- vis_sdr.chm11116070
ms.prod: visio
api_name:
- Visio.Selection.AddToGroup
ms.assetid: 8bef7960-271c-245d-dec0-eeea4af66097
ms.date: 06/08/2017
---


# Selection.AddToGroup Method (Visio)

Adds the selected shapes to the selected group.


## Syntax

 _expression_ . **AddToGroup**

 _expression_ A variable that represents a **Selection** object.


### Return Value

Nothing


## Remarks

The current selection must contain both the shapes to add and the group to which you want to add them. The group must be the primary selection or the only group in the selection.


## Example

The following macro shows how to use the  **AddToGroup** method to add selected shapes to a selected group.

Before running this macro, open the  **Basic Shapes** stencil or a document based on the **Basic Diagram** template.




```vb
 
Public Sub AddToGroup_Example() 
 
 Application.ActiveWindow.Page.Drop Application.Documents.Item("BASIC_U.VSS").Masters.ItemU("Square"), 3, 8 
 
 Application.ActiveWindow.Page.Drop Application.Documents.Item("BASIC_U.VSS").Masters.ItemU("Pentagon"), 4, 8 
 
 Application.ActiveWindow.SelectAll 
 
 ActiveWindow.DeselectAll 
 ActiveWindow.Select Application.ActiveWindow.Page.Shapes.ItemU("Pentagon"), visSelect 
 ActiveWindow.Select Application.ActiveWindow.Page.Shapes.ItemU("Square"), visSelect 
 ActiveWindow.Selection.Group 
 
 Application.ActiveWindow.Page.Drop Application.Documents.Item("BASIC_U.VSS").Masters.ItemU("Ellipse"), 5, 6 
 
 ActiveWindow.DeselectAll 
 ActiveWindow.Select Application.ActiveWindow.Page.Shapes.ItemU("Ellipse"), visSelect 
 ActiveWindow.Select Application.ActiveWindow.Page.Shapes.ItemU("Sheet.3"), visSelect 
 ActiveWindow.Selection.AddToGroup 
 
End Sub
```


