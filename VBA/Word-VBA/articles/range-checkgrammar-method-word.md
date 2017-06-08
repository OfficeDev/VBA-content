---
title: Range.CheckGrammar Method (Word)
keywords: vbawd10.chm157155532
f1_keywords:
- vbawd10.chm157155532
ms.prod: word
api_name:
- Word.Range.CheckGrammar
ms.assetid: 3ae0e80f-0165-be96-af12-b231d1f3a1b4
ms.date: 06/08/2017
---


# Range.CheckGrammar Method (Word)

Begins a spelling and grammar check for the specified range.


## Syntax

 _expression_ . **CheckGrammar**

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


## Remarks

If the range contains errors, this method displays the  **Spelling and Grammar** dialog box, with the **Check grammar** check box selected.


## Example

This example begins a spelling and grammar check on section two in MyDocument.doc.


```vb
Set Range2 = Documents("MyDocument.doc").Sections(2).Range 
Range2.CheckGrammar
```

This example begins a spelling and grammar check on the selection.




```
Selection.Range.CheckGrammar
```


## See also


#### Concepts


[Range Object](range-object-word.md)

