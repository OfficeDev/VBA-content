---
title: Worksheet.Names Property (Excel)
keywords: vbaxl10.chm175107
f1_keywords:
- vbaxl10.chm175107
ms.prod: excel
api_name:
- Excel.Worksheet.Names
ms.assetid: 4bdccfa9-7aa1-c3d6-6a89-5ce24aad2ad2
ms.date: 06/08/2017
---


# Worksheet.Names Property (Excel)

Returns a  **[Names](names-object-excel.md)** collection that represents all the worksheet-specific names (names defined with the "WorksheetName!" prefix). Read-only **Names** object.


## Syntax

 _expression_ . **Names**

 _expression_ A variable that represents a **Worksheet** object.


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


[Worksheet Object](worksheet-object-excel.md)

