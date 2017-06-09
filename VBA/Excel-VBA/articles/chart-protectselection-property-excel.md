---
title: Chart.ProtectSelection Property (Excel)
keywords: vbaxl10.chm149160
f1_keywords:
- vbaxl10.chm149160
ms.prod: excel
api_name:
- Excel.Chart.ProtectSelection
ms.assetid: a1b9cf7e-8cc3-f9fe-dfcf-c66469741edb
ms.date: 06/08/2017
---


# Chart.ProtectSelection Property (Excel)

 **True** if chart elements cannot be selected. Read/write **Boolean** .


## Syntax

 _expression_ . **ProtectSelection**

 _expression_ A variable that represents a **Chart** object.


## Remarks

When this property is  **True** , shapes cannot be added to the chart, and the **Click** and **DoubleClick** events for chart elements don't occur.

This property is not persisted when the file is saved. If you set this property to  **True** and then reopen the file, it will no longer be set to **True** .


## Example

This example prevents chart elements from being selected on embedded chart one on worksheet one.


```vb
Worksheets(1).ChartObjects(1).Chart.ProtectSelection = True
```


## See also


#### Concepts


[Chart Object](chart-object-excel.md)

