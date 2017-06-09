---
title: CanvasShapes.AddShape Method (Word)
keywords: vbawd10.chm7536657
f1_keywords:
- vbawd10.chm7536657
ms.prod: word
api_name:
- Word.CanvasShapes.AddShape
ms.assetid: b23c69f1-8653-a98f-d7f4-6648e0e214fa
ms.date: 06/08/2017
---


# CanvasShapes.AddShape Method (Word)

Adds an AutoShape to a drawing canvas. Returns a  **[Shape](shape-object-word.md)** object that represents the AutoShape.


## Syntax

 _expression_ . **AddShape**( **_Type_** , **_Left_** , **_Top_** , **_Width_** , **_Height_** )

 _expression_ Required. A variable that represents a **[CanvasShapes](canvasshapes-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **Long**|The type of shape to be returned. Can be any  **MsoAutoShape** constant.|
| _Left_|Required| **Single**|The position, measured in points, of the left edge of the AutoShape.|
| _Top_|Required| **Single**|The position, measured in points, of the top edge of the AutoShape.|
| _Width_|Required| **Single**|The width, measured in points, of the AutoShape.|
| _Height_|Required| **Single**|The height, measured in points, of the AutoShape.|

## Remarks

To change the type of an AutoShape that you've added, set the  **AutoShapeType** property.


## Example

This example creates a new canvas in the active document and adds a circle to the canvas.


```vb
Sub NewCanvasShape() 
 Dim shpCanvas As Shape 
 Dim shpCanvasShape As Shape 
 
 'Add a new drawing canvas to the active document 
 Set shpCanvas = ActiveDocument.Shapes.AddCanvas( _ 
 Left:=100, Top:=75, Width:=150, Height:=200) 
 
 'Add a circle to the drawing canvas 
 Set shpCanvasShape = shpCanvas.CanvasItems.AddShape( _ 
 Type:=msoShapeOval, Left:=25, Top:=25, _ 
 Width:=150, Height:=150) 
End Sub
```


## See also


#### Concepts


[CanvasShapes Collection](canvasshapes-object-word.md)

