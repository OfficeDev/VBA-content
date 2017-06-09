---
title: Range.Revisions Property (Word)
keywords: vbawd10.chm157155478
f1_keywords:
- vbawd10.chm157155478
ms.prod: word
api_name:
- Word.Range.Revisions
ms.assetid: cf71b684-991a-fb6d-09bc-eeecb16edec5
ms.date: 06/08/2017
---


# Range.Revisions Property (Word)

Returns a  **Revisions** collection that represents the tracked changes in the range. Read-only.


## Syntax

 _expression_ . **Revisions**

 _expression_ A variable that represents a **[Range](range-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning a Single Object from a Collection](http://msdn.microsoft.com/library/8c0b84c0-582b-32f7-68e0-6383d0661e74%28Office.15%29.aspx).


## Example

This example displays the number of tracked changes in the first section in the active document.


```vb
MsgBox ActiveDocument.Sections(1).Range.Revisions.Count
```

This example accepts all tracked changes in the first paragraph in the selection.




```vb
Set myRange = Selection.Paragraphs(1).Range 
myRange.Revisions.AcceptAll
```


## See also


#### Concepts


[Range Object](range-object-word.md)

