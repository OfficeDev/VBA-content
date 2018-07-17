---
title: Range.Conflicts Property (Word)
keywords: vbawd10.chm157155834
f1_keywords:
- vbawd10.chm157155834
ms.prod: word
api_name:
- Word.Range.Conflicts
ms.assetid: 908b36ff-a87a-255c-2b5d-e47dd6489bf7
ms.date: 06/08/2017
---


# Range.Conflicts Property (Word)

Returns a [Conflicts](conflicts-object-word.md) collection object that contains all the conflict objects in the range. Read-only.


## Syntax

 _expression_ . **Conflicts**

 _expression_ An expression that returns a **Range** object.


## Remarks

Use the  **Conflicts** property to return the[Conflicts](conflicts-object-word.md) collection object for a document. Use Conflicts( _Index_ ), where _Index_ is the conflict index number, to return a single[Conflict](conflict-object-word.md) object.


 **Note**  This property is only available for co authoring enabled documents. If you attempt to access this property on a document that is not enabled for co authoring, you will receive a run-time error.


## Example

The following code example displays the number of conflicts in the first paragraph of the active document.


```vb
MsgBox ActiveDocument.Paragraphs(1).Range.Conflicts.Count
```


## See also


#### Concepts


[Range Object](range-object-word.md)

