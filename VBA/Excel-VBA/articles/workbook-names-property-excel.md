---
title: Workbook.Names Property (Excel)
keywords: vbaxl10.chm199115
f1_keywords:
- vbaxl10.chm199115
ms.prod: excel
api_name:
- Excel.Workbook.Names
ms.assetid: 26be56ec-ea12-1600-602a-eb338d4a5a8b
ms.date: 06/08/2017
---


# Workbook.Names Property (Excel)

Returns a  **[Names](names-object-excel.md)** collection that represents all the names in the specified workbook (including all worksheet-specific names). Read-only **Names** object.


## Syntax

 _expression_ . **Names**

 _expression_ A variable that represents a **Workbook** object.


## Remarks

Using this property without an object qualifier is equivalent to using  `ActiveWorkbook.Names`.


## Example

This example defines the name "myName" for cell A1 on Sheet1.


```vb
ActiveWorkbook.Names.Add Name:="myName", RefersToR1C1:= _ 
 "=Sheet1!R1C1"
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

