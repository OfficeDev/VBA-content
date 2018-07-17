---
title: WorksheetFunction.Clean Method (Excel)
keywords: vbaxl10.chm137136
f1_keywords:
- vbaxl10.chm137136
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Clean
ms.assetid: ac5de21a-b087-ebd7-764b-1644475cd2a9
ms.date: 06/08/2017
---


# WorksheetFunction.Clean Method (Excel)

Removes all nonprintable characters from text.


## Syntax

 _expression_ . **Clean**( **_Arg1_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **String**|Any worksheet information from which you want to remove nonprintable characters.|

### Return Value

String


## Remarks

Use Clean on text imported from other applications that contains characters that may not print with your operating system. For example, you can use Clean to remove some low-level computer code that is frequently at the beginning and end of data files and cannot be printed.


 **Important**  The Clean function was designed to remove the first 32 nonprinting characters in the 7-bit ASCII code (values 0 through 31) from text. In the Unicode character set, there are additional nonprinting characters (values 127, 129, 141, 143, 144, and 157). By itself, the Clean function does not remove these additional nonprinting characters.


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

