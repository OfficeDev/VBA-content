---
title: Shape.TextFrame Property (Excel)
keywords: vbaxl10.chm636090
f1_keywords:
- vbaxl10.chm636090
ms.prod: excel
api_name:
- Excel.Shape.TextFrame
ms.assetid: cc2fbe92-e0c4-f0d5-52a3-a675d4baf573
ms.date: 06/08/2017
---


# Shape.TextFrame Property (Excel)

Returns a  **[TextFrame](textframe-object-excel.md)** object that contains the alignment and anchoring properties for the specified shape. Read-only.


## Syntax

 _expression_ . **TextFrame**

 _expression_ A variable that represents a **Shape** object.


## Example

This example causes text in the text frame in shape one to be justified. If shape one doesn't have a text frame, this example fails.


```vb
Worksheets(1).Shapes(1).TextFrame _ 
 .HorizontalAlignment = xlHAlignJustify
```


## See also


#### Concepts


[Shape Object](shape-object-excel.md)

