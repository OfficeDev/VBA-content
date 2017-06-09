---
title: ReplacementList Property
keywords: vbagr10.chm3077085
f1_keywords:
- vbagr10.chm3077085
ms.prod: excel
api_name:
- Excel.ReplacementList
ms.assetid: 14209e45-f0e9-a166-7970-ecf3ca79e570
ms.date: 06/08/2017
---


# ReplacementList Property

Returns the array of AutoCorrect replacements.

 _expression_. **ReplacementList**( **_Index_**)

 _expression_ Required. An expression that returns an **[AutoCorrect](autocorrect-object.md)** object.

 **Index** Optional **Variant**. The row index of the array of AutoCorrect replacements to be returned. The row is returned as a one-dimensional array with two elements: The first element is the text in column 1, and the second element is the text in column 2.

## Remarks

Use the  **[AddReplacement](addreplacement-method.md)** method to add an entry to the replacement list.


## Example

This example searches the replacement list for "Temperature" and displays the replacement entry if it exists.


```vb
repl = Application.AutoCorrect.ReplacementList 
For x = 1 To UBound(repl) 
 If repl(x, 1) = "Temperature" Then MsgBox repl(x, 2) 
Next
```


