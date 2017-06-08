---
title: WrapFormat.Type Property (Publisher)
keywords: vbapb10.chm786435
f1_keywords:
- vbapb10.chm786435
ms.prod: publisher
api_name:
- Publisher.WrapFormat.Type
ms.assetid: da53302c-ae95-5aa9-a4ce-32647a2569d6
ms.date: 06/08/2017
---


# WrapFormat.Type Property (Publisher)

Specifies how text wraps around the specified shape. Read/write.


## Syntax

 _expression_. **Type**

 _expression_A variable that represents a  **WrapFormat** object.


## Remarks

The  **Type** property value can be one of the **[PbWrapType](pbwraptype-enumeration-publisher.md)** constants declared in the Microsoft Publisher type library.


## Example

The following example adds an oval to the active publication and specifies that the publication text wrap around both the left and right sides of the square that surrounds the oval.


```vb
Sub SetTextWrapType() 
 Dim shpOval As Shape 
 
 Set shpOval = ActiveDocument.Pages(1).Shapes.AddShape( _ 
 Type:=msoShapeOval, Left:=36, Top:=36, _ 
 Width:=100, Height:=35) 
 
 With shpOval.TextWrap 
 .Type = pbWrapTypeSquare 
 .Side = pbWrapSideBoth 
 End With 
End Sub
```


