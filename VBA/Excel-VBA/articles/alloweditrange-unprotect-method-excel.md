---
title: AllowEditRange.Unprotect Method (Excel)
keywords: vbaxl10.chm725077
f1_keywords:
- vbaxl10.chm725077
ms.prod: excel
api_name:
- Excel.AllowEditRange.Unprotect
ms.assetid: 3c7679c6-828d-e1c4-7009-f42bad1a66d6
ms.date: 06/08/2017
---


# AllowEditRange.Unprotect Method (Excel)

Removes protection from a sheet or workbook. This method has no effect if the sheet or workbook isn't protected.


## Syntax

 _expression_ . **Unprotect**( **_Password_** )

 _expression_ A variable that represents an **AllowEditRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Password_|Optional| **Variant**|A string that denotes the case-sensitive password to use to unprotect the range of cells. If the range isn't protected with a password, this argument is ignored.|

## Remarks

If you forget the password, you cannot unprotect the sheet or workbook. It's a good idea to keep a list of your passwords and their corresponding document names in a safe place.


## See also


#### Concepts


[AllowEditRange Object](alloweditrange-object-excel.md)

