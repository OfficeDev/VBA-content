---
title: WorksheetFunction.Phonetic Method (Excel)
keywords: vbaxl10.chm137248
f1_keywords:
- vbaxl10.chm137248
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Phonetic
ms.assetid: a1da7aa0-f913-e64b-8863-212f8a4e261d
ms.date: 06/08/2017
---


# WorksheetFunction.Phonetic Method (Excel)

Extracts the phonetic (furigana) characters from a text string.


## Syntax

 _expression_ . **Phonetic**( **_Arg1_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Range**|Reference - a text string or a reference to a single cell or a range of cells that contain a furigana text string.|

### Return Value

String


## Remarks




- If reference is a range of cells, the furigana text string in the upper-left corner cell of the range is returned.
    
- If the reference is a range of nonadjacent cells, the #N/A error value is returned. 
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

