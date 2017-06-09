---
title: Selection.Words Property (Word)
keywords: vbawd10.chm158662707
f1_keywords:
- vbawd10.chm158662707
ms.prod: word
api_name:
- Word.Selection.Words
ms.assetid: bbbc7c5f-ce5a-2608-ba0c-e9769bff287a
ms.date: 06/08/2017
---


# Selection.Words Property (Word)

Returns a  **[Words](words-object-word.md)** collection that represents all the words in a selection. Read-only.


## Syntax

 _expression_ . **Words**

 _expression_ A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

Punctuation and paragraph marks in a document are included in the  **[Words](words-object-word.md)** collection. For information about returning a single member of a collection, see[Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example displays the number of words in the selection. Paragraphs marks, partial words, and punctuation are included in the count.


```vb
MsgBox "There are " &; Selection.Words.Count &; " words."
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

