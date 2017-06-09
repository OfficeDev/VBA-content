---
title: LayoutGuides.Rows Property (Publisher)
keywords: vbapb10.chm1114120
f1_keywords:
- vbapb10.chm1114120
ms.prod: publisher
api_name:
- Publisher.LayoutGuides.Rows
ms.assetid: a42286ef-d955-c39d-49a4-b0e54b4d1cec
ms.date: 06/08/2017
---


# LayoutGuides.Rows Property (Publisher)

Sets or returns a  **Long** that represents the number of rows in a layout guide. Read/write.


## Syntax

 _expression_. **Rows**

 _expression_A variable that represents a  **LayoutGuides** object.


## Example

This example sets the columns and rows for the layout guides.


```vb
Sub SetLayoutGuides() 
 With ActiveDocument.LayoutGuides 
 .Columns 
 .Rows 
 End With 
End Sub
```


