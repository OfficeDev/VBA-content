---
title: Range.Locks Property (Word)
keywords: vbawd10.chm157155832
f1_keywords:
- vbawd10.chm157155832
ms.prod: word
api_name:
- Word.Range.Locks
ms.assetid: 102673f2-8cb0-d235-c158-c65759592d56
ms.date: 06/08/2017
---


# Range.Locks Property (Word)

Returns a  **[CoAuthLocks](coauthlocks-object-word.md)** collection object that represents all the locks in the range. Read-only.


## Syntax

 _expression_ . **Locks**

 _expression_ An expression that returns a **[Range](range-object-word.md)** object.


## Remarks

Use the  **Locks** property to return the **[CoAuthLocks](coauthlocks-object-word.md)** collection.


 **Note**  This property is only available for co authoring enabled documents. If you attempt to access this property on a document that is not enabled for co authoring, you will receive a run-time error.


## Example

The following code example displays the number of locks in the first paragraph of the active document.


```vb
MsgBox ActiveDocument.Paragraphs(1).Range.Locks.Count
```


## See also


#### Concepts


[Range Object](range-object-word.md)

