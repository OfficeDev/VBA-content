---
title: WrapFormat.Side Property (Publisher)
keywords: vbapb10.chm786436
f1_keywords:
- vbapb10.chm786436
ms.prod: publisher
api_name:
- Publisher.WrapFormat.Side
ms.assetid: b7998643-216a-a294-bbee-e5f1947400a7
ms.date: 06/08/2017
---


# WrapFormat.Side Property (Publisher)

Returns or sets a  **PbWrapSideType** constant that indicates whether text should wrap around a shape. Read/write.


## Syntax

 _expression_. **Side**

 _expression_A variable that represents a  **WrapFormat** object.


### Return Value

PbWrapSideType


## Remarks

The  **Side** property value can be one of the **[PbWrapSideType](pbwrapsidetype-enumeration-publisher.md)** constants declared in the Microsoft Publisher type library.


## Example

This example adds an oval to the first page of the active publication and specifies that text wrap around both the left and right sides of the oval.


```vb
Sub SetTextWrapFormatProperties() 
 With ActiveDocument.Pages(1).Shapes.AddShape(Type:=msoShapeOval, _ 
 Left:=36, Top:=36, Width:=100, Height:=35) 
 With .TextWrap 
 .Type = pbWrapTypeSquare 
 .Side = pbWrapSideBoth 
 End With 
 End With 
End Sub
```


