---
title: Application.ExtendList Property (Excel)
keywords: vbaxl10.chm133243
f1_keywords:
- vbaxl10.chm133243
ms.prod: excel
api_name:
- Excel.Application.ExtendList
ms.assetid: b368047b-9d30-5a6f-a7db-748e3e91a3c0
ms.date: 06/08/2017
---


# Application.ExtendList Property (Excel)

 **True** if Microsoft Excel automatically extends formatting and formulas to new data that is added to a list. Read/write **Boolean** .


## Syntax

 _expression_ . **ExtendList**

 _expression_ A variable that represents an **Application** object.


## Remarks

To be extended, formats and formulas must appear in at least three of the five list rows or columns preceding the new row or column, and you must add the data to the bottom or to the right-hand side of the list.


## Example

This example sets Excel to not apply formatting and formulas to data subsequently added to an existing list.


```vb
Application.ExtendList = False
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

