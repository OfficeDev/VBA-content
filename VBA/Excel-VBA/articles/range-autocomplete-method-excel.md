---
title: Range.AutoComplete Method (Excel)
keywords: vbaxl10.chm144082
f1_keywords:
- vbaxl10.chm144082
ms.prod: excel
api_name:
- Excel.Range.AutoComplete
ms.assetid: 723a452f-34e1-fcd1-a2d6-4932c5cc0542
ms.date: 06/08/2017
---


# Range.AutoComplete Method (Excel)

Returns an AutoComplete match from the list. If there?s no AutoComplete match or if more than one entry in the list matches the string to complete, this method returns an empty string.


## Syntax

 _expression_ . **AutoComplete**( **_String_** )

 _expression_ A variable that represents a **Range** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _String_|Required| **String**|The string to complete.|

### Return Value

String


## Remarks

This method works even if the AutoComplete feature is disabled.


## Example

This example returns the AutoComplete match for the string segment ?Ap." An AutoComplete match is made if the column containing cell A5 contains a contiguous list and one of the entries in the list contains a match for the string.


```vb
s = Worksheets(1).Range("A5").AutoComplete("Ap") 
If Len(s) > 0 Then 
 MsgBox "Completes to " &; s 
Else 
 MsgBox "Has no completion" 
End If
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

