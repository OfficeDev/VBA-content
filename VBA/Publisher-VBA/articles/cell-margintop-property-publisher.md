---
title: Cell.MarginTop Property (Publisher)
keywords: vbapb10.chm5111829
f1_keywords:
- vbapb10.chm5111829
ms.prod: publisher
api_name:
- Publisher.Cell.MarginTop
ms.assetid: f408edd3-7199-b49a-817b-7b0e8461715c
ms.date: 06/08/2017
---


# Cell.MarginTop Property (Publisher)

Returns or sets a  **Variant** that represents the amount of space (in points) between the text and the top edge of a cell, text frame, or page. Read/write.


## Syntax

 _expression_. **MarginTop**

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


