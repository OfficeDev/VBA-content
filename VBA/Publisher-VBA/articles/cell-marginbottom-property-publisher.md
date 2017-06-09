---
title: Cell.MarginBottom Property (Publisher)
keywords: vbapb10.chm5111826
f1_keywords:
- vbapb10.chm5111826
ms.prod: publisher
api_name:
- Publisher.Cell.MarginBottom
ms.assetid: a05fd3a4-f4d5-232a-1f5d-0fa1bce136bd
ms.date: 06/08/2017
---


# Cell.MarginBottom Property (Publisher)

Returns or sets a  **Variant** that represents the amount of space (in points) between the text and the bottom edge of a cell, text frame, or page. Read/write.


## Syntax

 _expression_. **MarginBottom**

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


