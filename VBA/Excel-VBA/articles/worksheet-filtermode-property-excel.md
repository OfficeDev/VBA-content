---
title: Worksheet.FilterMode Property (Excel)
keywords: vbaxl10.chm175100
f1_keywords:
- vbaxl10.chm175100
ms.prod: excel
api_name:
- Excel.Worksheet.FilterMode
ms.assetid: d9bcaa8a-caf3-96a4-445d-d957a987b057
ms.date: 06/08/2017
---


# Worksheet.FilterMode Property (Excel)

 **True** if the worksheet is in the filter mode. Read-only **Boolean** .


## Syntax

 _expression_ . **FilterMode**

 _expression_ A variable that represents a **Worksheet** object.


## Example

In the following example, the code returns True if the worksheet is in the filter mode.


```vb
Dim Worksheet1 As Worksheet 
 
Dim returnValue As Boolean 
returnValue = Worksheet1.FilterMode
```


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

