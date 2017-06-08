---
title: Worksheet.ShowAllData Method (Excel)
keywords: vbaxl10.chm175126
f1_keywords:
- vbaxl10.chm175126
ms.prod: excel
api_name:
- Excel.Worksheet.ShowAllData
ms.assetid: 412acb6c-f83d-44d4-20b5-54a2b7c66284
ms.date: 06/08/2017
---


# Worksheet.ShowAllData Method (Excel)

Makes all rows of the currently filtered list visible. If AutoFilter is in use, this method changes the arrows to "All."


## Syntax

 _expression_ . **ShowAllData**

 _expression_ A variable that represents a **Worksheet** object.


## Example

This example makes all data on Sheet1 visible. The example should be run on a worksheet that contains a list you filtered using the  **AutoFilter** command.


```vb
Worksheets("Sheet1").ShowAllData
```


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

