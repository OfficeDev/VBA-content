---
title: Style.IncludeProtection Property (Excel)
keywords: vbaxl10.chm177085
f1_keywords:
- vbaxl10.chm177085
ms.prod: excel
api_name:
- Excel.Style.IncludeProtection
ms.assetid: 666afea1-4a2a-7f44-ecff-d9d44098a527
ms.date: 06/08/2017
---


# Style.IncludeProtection Property (Excel)

 **True** if the style includes the **FormulaHidden** and **Locked** protection properties. Read/write **Boolean** .


## Syntax

 _expression_ . **IncludeProtection**

 _expression_ A variable that represents a **Style** object.


## Example

This example sets the style attached to cell A1 on Sheet1 to include protection format.


```vb
Worksheets("Sheet1").Range("A1").Style.IncludeProtection = True
```


## See also


#### Concepts


[Style Object](style-object-excel.md)

