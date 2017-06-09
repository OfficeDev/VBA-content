---
title: Cell.MarginRight Property (Publisher)
keywords: vbapb10.chm5111828
f1_keywords:
- vbapb10.chm5111828
ms.prod: publisher
api_name:
- Publisher.Cell.MarginRight
ms.assetid: d297222e-7fc1-9225-e098-1a85d7734d77
ms.date: 06/08/2017
---


# Cell.MarginRight Property (Publisher)

Returns or sets a  **Variant** that represents the amount of space (in points) between the text and the right edge of a cell, text frame, or page. Read/write.


## Syntax

 _expression_. **MarginRight**

 _expression_A variable that represents a  **Cell** object.


## Example

This example sets the margins of the active publication to two inches.


```vb
Sub SetPageMargins() 
 
 With ActiveDocument.LayoutGuides 
 .MarginTop = Application.InchesToPoints(Value:=2) 
 .MarginBottom = Application.InchesToPoints(Value:=2) 
 .MarginLeft = Application.InchesToPoints(Value:=2) 
 .MarginRight = Application.InchesToPoints(Value:=2) 
 End With 
 
End Sub
```


