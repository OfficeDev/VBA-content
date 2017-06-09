---
title: WorksheetFunction.Trim Method (Excel)
keywords: vbaxl10.chm137126
f1_keywords:
- vbaxl10.chm137126
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Trim
ms.assetid: 1e596960-90d8-87f8-9f1f-3a5c9e302e0c
ms.date: 06/08/2017
---


# WorksheetFunction.Trim Method (Excel)

Removes all spaces from text except for single spaces between words. Use TRIM on text that you have received from another application that may have irregular spacing.


## Syntax

 _expression_ . **Trim**( **_Arg1_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **String**|Text - the text from which you want spaces removed.|

### Return Value

String


## Remarks


 **Important**  The TRIM function was designed to trim the 7-bit ASCII space character (value 32) from text. In the Unicode character set, there is an additional space character called the nonbreaking space character that has a decimal value of 160. This character is commonly used in Web pages as the HTML entity,  **&;nbsp;** . By itself, the TRIM function does not remove this nonbreaking space character.


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

