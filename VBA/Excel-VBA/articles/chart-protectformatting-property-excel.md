---
title: Chart.ProtectFormatting Property (Excel)
keywords: vbaxl10.chm149157
f1_keywords:
- vbaxl10.chm149157
ms.prod: excel
api_name:
- Excel.Chart.ProtectFormatting
ms.assetid: 71630b7f-6c89-869d-cd5b-d0a7bacd904a
ms.date: 06/08/2017
---


# Chart.ProtectFormatting Property (Excel)

 **True** if chart formatting cannot be modified by the user. Read/write **Boolean** .


## Syntax

 _expression_ . **ProtectFormatting**

 _expression_ A variable that represents a **Chart** object.


## Remarks

This property is not persisted when the file is saved. If you set this property to  **True** and then reopen the file, it will no longer be set to **True** .


## Example

This example protects the formatting of embedded chart one on worksheet one.


```vb
Worksheets(1).ChartObjects(1).Chart.ProtectFormatting = True
```


## See also


#### Concepts


[Chart Object](chart-object-excel.md)

