---
title: Range.InsertAutoText Method (Word)
keywords: vbawd10.chm157155511
f1_keywords:
- vbawd10.chm157155511
ms.prod: word
api_name:
- Word.Range.InsertAutoText
ms.assetid: d87ae18c-e527-bcf4-4939-5512a6fdaaf5
ms.date: 06/08/2017
---


# Range.InsertAutoText Method (Word)

Attempts to match the text in the specified range or the text surrounding the range with an existing AutoText entry name.


## Syntax

 _expression_ . **InsertAutoText**

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


## Remarks

If Word finds a match,  **InsertAutoText** inserts the AutoText entry to replace that text. If Word cannot find a match, an error occurs.

You can use the  **Insert** method with an **AutoTextEntry** object to insert a specific AutoText entry.


## Example

This example inserts an AutoText entry that matches the text around a selection.


```
Documents.Add 
Selection.TypeText "Best w" 
Selection.Range.InsertAutoText
```

This example inserts an AutoText entry with a name that matches the first word in the active document.




```vb
Documents.Add 
Selection.TypeText "In " 
Set myRange = ActiveDocument.Words(1) 
myRange.InsertAutoText
```


## See also


#### Concepts


[Range Object](range-object-word.md)

