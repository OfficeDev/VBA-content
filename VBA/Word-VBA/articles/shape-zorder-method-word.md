---
title: Shape.ZOrder Method (Word)
keywords: vbawd10.chm161480728
f1_keywords:
- vbawd10.chm161480728
ms.prod: word
api_name:
- Word.Shape.ZOrder
ms.assetid: b6729719-44b0-a069-0cbe-b694b88ab65a
ms.date: 06/08/2017
---


# Shape.ZOrder Method (Word)

Moves the specified shape in front of or behind other shapes in the collection (that is, changes the shape's position in the z-order).


## Syntax

 _expression_ . **ZOrder**( **_ZOrderCmd_** )

 _expression_ An expression that returns a **[Shape](shape-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ZOrderCmd_|Required| **MsoZOrderCmd**|Specifies where to move the specified shape relative to the other shapes.|

### Return Value

Nothing


## Remarks

Use the  **[ZOrderPosition](shape-zorderposition-property-word.md)** property to determine a shape's current position in the z-order.


## Example

This example adds an oval to the active document and then places the oval as second from the back in the z-order if there is at least one other shape on the document.


```vb
With ActiveDocument.Shapes.AddShape(Type:=msoShapeOval, Left:=100, _ 
 Top:=100, Width:=100, Height:=300) 
 While .ZOrderPosition > 2 
 .ZOrder msoSendBackward 
 Wend 
End With
```


## See also


#### Concepts


[Shape Object](shape-object-word.md)

