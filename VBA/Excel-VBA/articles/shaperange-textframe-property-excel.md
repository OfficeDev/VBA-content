---
title: ShapeRange.TextFrame Property (Excel)
keywords: vbaxl10.chm640097
f1_keywords:
- vbaxl10.chm640097
ms.prod: excel
api_name:
- Excel.ShapeRange.TextFrame
ms.assetid: b72b9c3e-c41c-dce9-46ba-ee156ba52676
ms.date: 06/08/2017
---


# ShapeRange.TextFrame Property (Excel)

Returns a  **[TextFrame](textframe-object-excel.md)** object that contains the alignment and anchoring properties for the specified shape. Read-only.


## Syntax

 _expression_ . **TextFrame**

 _expression_ A variable that represents a **ShapeRange** object.


## Example

This example causes text in the text frame in shape one to be justified. If shape one doesn't have a text frame, this example fails.


```vb
Worksheets(1).Shapes(1).TextFrame _ 
 .HorizontalAlignment = xlHAlignJustify
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-excel.md)

