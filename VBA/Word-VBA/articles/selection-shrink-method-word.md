---
title: Selection.Shrink Method (Word)
keywords: vbawd10.chm158662957
f1_keywords:
- vbawd10.chm158662957
ms.prod: word
api_name:
- Word.Selection.Shrink
ms.assetid: ed364c95-3b9d-44dc-b120-db23aedfeaed
ms.date: 06/08/2017
---


# Selection.Shrink Method (Word)

Shrinks the selection to the next smaller unit of text.


## Syntax

 _expression_ . **Shrink**

 _expression_ A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

The unit progression for this method is as follows: entire document, section, paragraph, sentence, word, insertion point.


## Example

This example collapses the selected text to the next smaller unit of text.


```vb
If Selection.Type = wdSelectionNormal Then 
 Selection.Shrink 
Else 
 MsgBox "You need to select some text." 
End If
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

