---
title: ListFormat.SingleList Property (Word)
keywords: vbawd10.chm163577928
f1_keywords:
- vbawd10.chm163577928
ms.prod: word
api_name:
- Word.ListFormat.SingleList
ms.assetid: b2ec4d04-bc2b-b369-b213-f7e25ca894a4
ms.date: 06/08/2017
---


# ListFormat.SingleList Property (Word)

 **True** if the specified **ListFormat** object contains only one list. Read-only **Boolean** .


## Syntax

 _expression_ . **SingleList**

 _expression_ An expression that returns a **[ListFormat](listformat-object-word.md)** object.


## Example

This example checks the selection to see whether it only contains one list. If it does, the example applies the default numbered list template to the selection.


```vb
temp = Selection.Range.ListFormat.SingleList 
If temp = True Then 
 Selection.Range.ListFormat.ApplyNumberDefault 
End If
```


## See also


#### Concepts


[ListFormat Object](listformat-object-word.md)

