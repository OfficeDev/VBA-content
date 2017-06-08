---
title: Chart.ProtectData Property (Excel)
keywords: vbaxl10.chm149158
f1_keywords:
- vbaxl10.chm149158
ms.prod: excel
api_name:
- Excel.Chart.ProtectData
ms.assetid: 29eb3e29-6005-70bd-cb38-053a5d54ed96
ms.date: 06/08/2017
---


# Chart.ProtectData Property (Excel)

 **True** if series formulas cannot be modified by the user. Read/write **Boolean** .


## Syntax

 _expression_ . **ProtectData**

 _expression_ A variable that represents a **Chart** object.


## Remarks

This property is not persisted when the file is saved. If you set this property to  **True** and then reopen the file, it will no longer be set to **True** .


## Example

This example protects the data on embedded chart one on worksheet one.


```vb
Worksheets(1).ChartObjects(1).Chart.ProtectData = True
```


## See also


#### Concepts


[Chart Object](chart-object-excel.md)

