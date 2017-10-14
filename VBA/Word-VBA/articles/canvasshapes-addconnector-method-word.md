---
title: CanvasShapes.AddConnector Method (Word)
keywords: vbawd10.chm7536651
f1_keywords:
- vbawd10.chm7536651
ms.prod: word
api_name:
- Word.CanvasShapes.AddConnector
ms.assetid: 0bfb15ae-0d0d-1864-bdf3-941e22090696
ms.date: 06/08/2017
---


# CanvasShapes.AddConnector Method (Word)

Returns a  **[Shape](shape-object-word.md)** object that represents a connecting line between two shapes in a drawing canvas.


## Syntax

 _expression_ . **AddConnector**( **_Type_** , **_BeginX_** , **_BeginY_** , **_EndX_** , **_EndY_** )

 _expression_ Required. A variable that represents a **[CanvasShapes](canvasshapes-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **MsoConnectorType**|The type of connector.|
| _BeginX_|Required| **Single**|The horizontal position that marks the beginning of the connector.|
| _BeginY_|Required| **Single**|The vertical position that marks the beginning of the connector.|
| _EndX_|Required| **Single**|The horizontal position that marks the end of the connector.|
| _EndY_|Required| **Single**|The vertical position that marks the end of the connector.|

## Example

The following example adds a straight connector to a new canvas in a new document.


```vb
Sub AddCanvasConnector() 
 
 Dim docNew As Document 
 Dim shpCanvas As Shape 
 
 Set docNew = Documents.Add 
 
 'Add drawing canvas to new document 
 Set shpCanvas = docNew.Shapes.AddCanvas( _ 
 Left:=150, Top:=150, Width:=200, Height:=300) 
 
 'Add connector to the drawing canvas 
 shpCanvas.CanvasItems.AddConnector _ 
 Type:=msoConnectorStraight, BeginX:=150, _ 
 BeginY:=150, EndX:=200, EndY:=200 
 
End Sub
```


## See also


#### Concepts


[CanvasShapes Collection](canvasshapes-object-word.md)

