---
title: WorksheetFunction.Substitute Method (Excel)
keywords: vbaxl10.chm137128
f1_keywords:
- vbaxl10.chm137128
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Substitute
ms.assetid: 1e02eb86-6902-0073-33ea-8d9f08b4eb14
ms.date: 06/08/2017
---


# WorksheetFunction.Substitute Method (Excel)

Substitutes new_text for old_text in a text string. Use SUBSTITUTE when you want to replace specific text in a text string; use REPLACE when you want to replace any text that occurs in a specific location in a text string.


## Syntax

 _expression_ . **Substitute**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **String**|Text - the text or the reference to a cell containing text for which you want to substitute characters.|
| _Arg2_|Required| **String**|Old_text - the text you want to replace.|
| _Arg3_|Required| **String**|New_text - the text you want to replace old_text with.|
| _Arg4_|Optional| **Variant**|Instance_num - specifies which occurrence of old_text you want to replace with new_text. If you specify instance_num, only that instance of old_text is replaced. Otherwise, every occurrence of old_text in text is changed to new_text.|

### Return Value

String


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

