---
title: Interior.Pattern Property (Excel)
keywords: vbaxl10.chm551076
f1_keywords:
- vbaxl10.chm551076
ms.prod: excel
api_name:
- Excel.Interior.Pattern
ms.assetid: 90587a6d-273c-00df-bb12-1a4415591705
ms.date: 06/08/2017
---


# Interior.Pattern Property (Excel)

Returns or sets a  **Variant** value, containing an **[xlPattern](xlpattern-enumeration-excel.md)** constant, that represents the interior pattern.


## Syntax

 _expression_ . **Pattern**

 _expression_ A variable that represents an **Interior** object.


## Example

This example adds a crisscross pattern to the interior of cell A1 on Sheet1.


```vb
Worksheets("Sheet1").Range("A1"). _ 
 Interior.Pattern = xlPatternCrissCross
```


## See also


#### Concepts


[Interior Object](interior-object-excel.md)

