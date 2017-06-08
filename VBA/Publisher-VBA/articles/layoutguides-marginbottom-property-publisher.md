---
title: LayoutGuides.MarginBottom Property (Publisher)
keywords: vbapb10.chm1114115
f1_keywords:
- vbapb10.chm1114115
ms.prod: publisher
api_name:
- Publisher.LayoutGuides.MarginBottom
ms.assetid: 9d11c4d9-8f53-7882-be40-200833a29fb6
ms.date: 06/08/2017
---


# LayoutGuides.MarginBottom Property (Publisher)

Returns or sets a  **Variant** that represents the amount of space (in points) between the text and the bottom edge of a cell, text frame, or page. Read/write.


## Syntax

 _expression_. **MarginBottom**

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


