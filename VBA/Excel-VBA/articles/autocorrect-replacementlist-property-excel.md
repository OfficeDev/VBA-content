---
title: AutoCorrect.ReplacementList Property (Excel)
keywords: vbaxl10.chm545076
f1_keywords:
- vbaxl10.chm545076
ms.prod: excel
api_name:
- Excel.AutoCorrect.ReplacementList
ms.assetid: 10bc895b-cd97-26a7-8b9e-4ac9347ebfc1
ms.date: 06/08/2017
---


# AutoCorrect.ReplacementList Property (Excel)

Returns the array of AutoCorrect replacements.


## Syntax

 _expression_ . **ReplacementList**( **_Index_** )

 _expression_ A variable that represents an **AutoCorrect** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|The row index of the array of AutoCorrect replacements to be returned. The row is returned as a one-dimensional array with two elements: The first element is the text in column 1, and the second element is the text in column 2.|

## Remarks

If  _Index_ is not specified, this method returns a two-dimensional array. Each row in the array contains one replacement, as shown in the following table.



|**Column**|**Contents**|
|:-----|:-----|
|1|The text to be replaced|
|2|The replacement text|
Use the  **[AddReplacement](autocorrect-addreplacement-method-excel.md)** method to add an entry to the replacement list.


## Example

This example searches the replacement list for "Temperature" and displays the replacement entry if it exists.


```vb
repl = Application.AutoCorrect.ReplacementList 
For x = 1 To UBound(repl) 
 If repl(x, 1) = "Temperature" Then MsgBox repl(x, 2) 
Next
```


## See also


#### Concepts


[AutoCorrect Object](autocorrect-object-excel.md)

