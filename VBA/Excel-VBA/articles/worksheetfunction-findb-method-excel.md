---
title: WorksheetFunction.FindB Method (Excel)
keywords: vbaxl10.chm137154
f1_keywords:
- vbaxl10.chm137154
ms.prod: excel
api_name:
- Excel.WorksheetFunction.FindB
ms.assetid: 463309cb-7747-6ee4-899b-677222e2dbda
ms.date: 06/08/2017
---


# WorksheetFunction.FindB Method (Excel)

FIND and FINDB locate one text string within a second text string, and return the number of the starting position of the first text string from the first character of the second text string. 


## Syntax

 _expression_ . **FindB**( **_Arg1_** , **_Arg2_** , **_Arg3_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **String**|Find_text - the text you want to find.|
| _Arg2_|Required| **String**|Within_text - the text containing the text you want to find.|
| _Arg3_|Optional| **Variant**|Start_num - specifies the character at which to start the search. The first character in within_text is character number 1. If you omit start_num, it is assumed to be 1.|

### Return Value

Double


## Remarks


 **Important**  FIND is intended for use with languages that use the single-byte character set (SBCS), whereas FINDB is intended for use with languages that use the double-byte character set (DBCS). The default language setting on your computer affects the return value in the following way:


- FIND always counts each character, whether single-byte or double-byte, as 1, no matter what the default language setting is.
    
-  FINDB counts each double-byte character as 2 when you have enabled the editing of a language that supports DBCS and then set it as the default language. Otherwise, FINDB counts each character as 1.
    
 The languages that support DBCS include Japanese, Chinese (Simplified), Chinese (Traditional), and Korean.


- FIND and FINDB are case sensitive and don't allow wildcard characters. If you don't want to do a case sensitive search or use wildcard characters, you can use SEARCH and SEARCHB.
    
- If find_text is "" (empty text), FIND matches the first character in the search string (that is, the character numbered start_num or 1).
    
- Find_text cannot contain any wildcard characters.
    
- If find_text does not appear in within_text, FIND and FINDB return the #VALUE! error value.
    
- If start_num is not greater than zero, FIND and FINDB return the #VALUE! error value.
    
- If start_num is greater than the length of within_text, FIND and FINDB return the #VALUE! error value.
    
- Use start_num to skip a specified number of characters. Using FIND as an example, suppose you are working with the text string "AYF0093.YoungMensApparel". To find the number of the first "Y" in the descriptive part of the text string, set start_num equal to 8 so that the serial-number portion of the text is not searched. FIND begins with character 8, finds find_text at the next character, and returns the number 9. FIND always returns the number of characters from the start of within_text, counting the characters you skip if start_num is greater than 1.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

