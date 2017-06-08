---
title: Workbook.Unprotect Method (Excel)
keywords: vbaxl10.chm199157
f1_keywords:
- vbaxl10.chm199157
ms.prod: excel
api_name:
- Excel.Workbook.Unprotect
ms.assetid: 39387902-a8a4-7bf2-44d7-c5bde6725778
ms.date: 06/08/2017
---


# Workbook.Unprotect Method (Excel)

Removes protection from a sheet or workbook. This method has no effect if the sheet or workbook isn't protected.


## Syntax

 _expression_ . **Unprotect**( **_Password_** )

 _expression_ A variable that represents a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Password_|Optional| **Variant**|A string that denotes the case-sensitive password to use to unprotect the sheet or workbook. If the sheet or workbook isn't protected with a password, this argument is ignored. If you omit this argument for a sheet that's protected with a password, you'll be prompted for the password. If you omit this argument for a workbook that's protected with a password, the method fails.|

## Remarks

If you forget the password, you cannot unprotect the sheet or workbook. It's a good idea to keep a list of your passwords and their corresponding document names in a safe place.


## Example

This example removes protection from the active workbook.


```vb
ActiveWorkbook.Unprotect
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

