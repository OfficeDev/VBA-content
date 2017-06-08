---
title: Border.Weight Property (Excel)
keywords: vbaxl10.chm547076
f1_keywords:
- vbaxl10.chm547076
ms.prod: excel
api_name:
- Excel.Border.Weight
ms.assetid: c6b9a812-60e6-245d-e86e-fb385581f890
ms.date: 06/08/2017
---


# Border.Weight Property (Excel)

Returns or sets a  **[XlBorderWeight](xlborderweight-enumeration-excel.md)** value that represents the weight of the border.


## Syntax

 _expression_ . **Weight**

 _expression_ A variable that represents a **Border** object.


## Example

This example sets the border weight for oval one on Sheet1.


```vb
Worksheets("Sheet1").Ovals(1).Border.Weight = xlMedium
```


## See also


#### Concepts


[Border Object](border-object-excel.md)

