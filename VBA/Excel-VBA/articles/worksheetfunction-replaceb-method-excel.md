---
title: WorksheetFunction.ReplaceB Method (Excel)
keywords: vbaxl10.chm137156
f1_keywords:
- vbaxl10.chm137156
ms.prod: excel
api_name:
- Excel.WorksheetFunction.ReplaceB
ms.assetid: 8853dcdd-6cd0-6ac5-1a71-27054f2a4776
ms.date: 06/08/2017
---


# WorksheetFunction.ReplaceB Method (Excel)

REPLACEB replaces part of a text string, based on the number of bytes you specify, with a different text string. 


## Syntax

 _expression_ . **ReplaceB**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **String**|Old_text - text in which you want to replace some characters.|
| _Arg2_|Required| **Double**|Start_num - the position of the character in old_text that you want to replace with new_text.|
| _Arg3_|Required| **Double**|Num_chars - the number of characters in old_text that you want REPLACE to replace with new_text.|
| _Arg4_|Required| **String**|New_text - the text that will replace characters in old_text.|

### Return Value

String


## Remarks


 **Important**  REPLACE is intended for use with languages that use the single-byte character set (SBCS), whereas REPLACEB is intended for use with languages that use the double-byte character set (DBCS). The default language setting on your computer affects the return value in the following way:


- REPLACE always counts each character, whether single-byte or double-byte, as 1, no matter what the default language setting is.
    
-  REPLACEB counts each double-byte character as 2 when you have enabled the editing of a language that supports DBCS and then set it as the default language. Otherwise, REPLACEB counts each character as 1.
    
The languages that support DBCS include Japanese, Chinese (Simplified), Chinese (Traditional), and Korean. 


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

