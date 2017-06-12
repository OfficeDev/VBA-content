---
title: WorksheetFunction.Roman Method (Excel)
keywords: vbaxl10.chm137245
f1_keywords:
- vbaxl10.chm137245
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Roman
ms.assetid: 59c27951-ad7e-7fe9-af5c-91ff1c68e7d4
ms.date: 06/08/2017
---


# WorksheetFunction.Roman Method (Excel)

Converts an arabic numeral to roman, as text.


## Syntax

 _expression_ . **Roman**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Number - the Arabic numeral you want converted.|
| _Arg2_|Optional| **Variant**|Form - a number specifying the type of roman numeral you want. The roman numeral style ranges from Classic to Simplified, becoming more concise as the value of form increases.|

### Return Value

String


## Remarks



|**Form**|**Type**|
|:-----|:-----|
|0 or omitted|Classic.|
|1|More concise. See example below.|
|2|More concise. See example below.|
|3|More concise. See example below.|
|4|Simplified.|
|TRUE|Classic.|
|FALSE|Simplified.|

- If number is negative, the #VALUE! error value is returned.
    
- If number is greater than 3999, the #VALUE! error value is returned.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

