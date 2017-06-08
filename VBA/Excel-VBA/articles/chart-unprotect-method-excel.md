---
title: Chart.Unprotect Method (Excel)
keywords: vbaxl10.chm148095
f1_keywords:
- vbaxl10.chm148095
ms.prod: excel
api_name:
- Excel.Chart.Unprotect
ms.assetid: 59a367bd-037b-84aa-5b2f-d532614ed347
ms.date: 06/08/2017
---


# Chart.Unprotect Method (Excel)

Removes protection from a sheet or workbook. This method has no effect if the sheet or workbook isn't protected.


## Syntax

 _expression_ . **Unprotect**( **_Password_** )

 _expression_ A variable that represents a **Chart** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Password_|Optional| **Variant**|A string that denotes the case-sensitive password to use to unprotect the chart. If the chart isn't protected with a password, this argument is ignored.|

## Remarks

If you forget the password, you cannot unprotect the sheet or workbook. It's a good idea to keep a list of your passwords and their corresponding document names in a safe place.


## See also


#### Concepts


[Chart Object](chart-object-excel.md)

