---
title: Range.Subdocuments Property (Word)
keywords: vbawd10.chm157155487
f1_keywords:
- vbawd10.chm157155487
ms.prod: word
api_name:
- Word.Range.Subdocuments
ms.assetid: c06afeb9-7e83-d858-d863-9582962c8254
ms.date: 06/08/2017
---


# Range.Subdocuments Property (Word)

Returns a  **Subdocuments** collection that represents all the subdocuments in the specified range or document. Read-only.


## Syntax

 _expression_ . **Subdocuments**

 _expression_ A variable that represents a **[Range](range-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example displays the number of subdocuments embedded in the active document.


```vb
MsgBox ActiveDocument.Range.Subdocuments.Count
```


## See also


#### Concepts


[Range Object](range-object-word.md)

