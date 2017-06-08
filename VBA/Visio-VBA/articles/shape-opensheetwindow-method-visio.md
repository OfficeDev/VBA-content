---
title: Shape.OpenSheetWindow Method (Visio)
keywords: vis_sdr.chm11216415
f1_keywords:
- vis_sdr.chm11216415
ms.prod: visio
api_name:
- Visio.Shape.OpenSheetWindow
ms.assetid: 744b72f5-381a-48fc-407f-20ffe815c54e
ms.date: 06/08/2017
---


# Shape.OpenSheetWindow Method (Visio)

Opens a ShapeSheet window for a  **Shape** object.


## Syntax

 _expression_ . **OpenSheetWindow**

 _expression_ A variable that represents a **Shape** object.


### Return Value

Window


## Remarks

The  **OpenSheetWindow** method opens a new ShapeSheet window for the shape even if the information is already displayed in another window.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **OpenSheetWindow** method to open the ShapeSheet window of a **Shape** object.


```vb
 
Public Sub OpenSheetWindow_Example() 
 
 Dim vsoShape As Visio.Shape 
 Dim vsoSheetWindow As Visio.Window 
 
 'Draw a shape. 
 Set vsoShape = ActivePage.DrawRectangle(1, 1, 2, 3) 
 
 'Open the ShapeSheet window of vsoShape. 
 Set vsoSheetWindow = vsoShape.OpenSheetWindow 
 
End Sub
```


