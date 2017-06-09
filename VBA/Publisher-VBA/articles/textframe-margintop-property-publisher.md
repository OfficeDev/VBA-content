---
title: TextFrame.MarginTop Property (Publisher)
keywords: vbapb10.chm3866645
f1_keywords:
- vbapb10.chm3866645
ms.prod: publisher
api_name:
- Publisher.TextFrame.MarginTop
ms.assetid: 9709eefe-0857-f228-aa56-780c4789a413
ms.date: 06/08/2017
---


# TextFrame.MarginTop Property (Publisher)

Returns or sets a  **Variant** that represents the amount of space (in points) between the text and the top edge of a cell, text frame, or page. Read/write.


## Syntax

 _expression_. **MarginTop**

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


