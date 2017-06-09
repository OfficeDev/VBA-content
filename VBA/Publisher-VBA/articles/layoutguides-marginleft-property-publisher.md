---
title: LayoutGuides.MarginLeft Property (Publisher)
keywords: vbapb10.chm1114116
f1_keywords:
- vbapb10.chm1114116
ms.prod: publisher
api_name:
- Publisher.LayoutGuides.MarginLeft
ms.assetid: 02d1a544-3e41-3875-3027-61bdc465e89b
ms.date: 06/08/2017
---


# LayoutGuides.MarginLeft Property (Publisher)

Returns or sets a  **Variant** that represents the amount of space (in points) between the text and the left edge of a cell, text frame, or page. Read/write.


## Syntax

 _expression_. **MarginLeft**

 _expression_A variable that represents a  **LayoutGuides** object.


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


