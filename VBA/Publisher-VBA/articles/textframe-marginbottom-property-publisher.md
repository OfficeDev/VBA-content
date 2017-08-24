---
title: TextFrame.MarginBottom Property (Publisher)
keywords: vbapb10.chm3866647
f1_keywords:
- vbapb10.chm3866647
ms.prod: publisher
api_name:
- Publisher.TextFrame.MarginBottom
ms.assetid: 55858bba-1103-48ba-64d6-5cc5ab677867
ms.date: 06/08/2017
---


# TextFrame.MarginBottom Property (Publisher)

Returns or sets a  **Variant** that represents the amount of space (in points) between the text and the bottom edge of a cell, text frame, or page. Read/write.


## Syntax

 _expression_. **MarginBottom**

 _expression_A variable that represents a  **TextFrame** object.


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


