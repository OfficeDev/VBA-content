---
title: Shapes.AddPolyline Method (Word)
keywords: vbawd10.chm161415184
f1_keywords:
- vbawd10.chm161415184
ms.prod: word
api_name:
- Word.Shapes.AddPolyline
ms.assetid: a9e25bec-095d-e088-246a-34b813ab0328
ms.date: 06/08/2017
---


# Shapes.AddPolyline Method (Word)

Adds an open or closed polygon to a drawing canvas. Returns a  **Shape** object that represents the polygon and adds it to the **CanvasShapes** collection.


## Syntax

 _expression_ . **AddPolyline**( **_SafeArrayOfPoints_** )

 _expression_ Required. A variable that represents a **[Shapes](shapes-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SafeArrayOfPoints_|Required| **Variant**|An array of coordinate pairs that specifies the polyline drawing's vertices.|

## Remarks

To form a closed polygon, assign the same coordinates to the first and last vertices in the polyline drawing.


## Example

This example creates a V-shaped open polyline in a new drawing canvas.


```vb
Sub NewCanvasPolyline() 
 Dim docNew As Document 
 Dim shpCanvas As Shape 
 Dim sngArray(1 To 3, 1 To 2) As Single 
 
 'Creates a new document and adds a drawing canvas 
 Set docNew = Documents.Add 
 Set shpCanvas = docNew.Shapes.AddCanvas( _ 
 Left:=100, Top:=75, Width:=200, Height:=300) 
 
 'Sets the coordinates of the array 
 sngArray(1, 1) = 100 
 sngArray(1, 2) = 75 
 sngArray(2, 1) = 150 
 sngArray(2, 2) = 100 
 sngArray(3, 1) = 100 
 sngArray(3, 2) = 125 
 
 'Adds a V-shaped open polyline to the drawing canvas 
 shpCanvas.CanvasItems.AddPolyline SafeArrayOfPoints:=sngArray 
End Sub
```


## See also


#### Concepts


[Shapes Collection Object](shapes-object-word.md)

