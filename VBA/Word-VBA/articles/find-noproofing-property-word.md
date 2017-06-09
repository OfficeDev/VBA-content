---
title: Find.NoProofing Property (Word)
keywords: vbawd10.chm162529314
f1_keywords:
- vbawd10.chm162529314
ms.prod: word
api_name:
- Word.Find.NoProofing
ms.assetid: 4e13dab9-8bff-5615-c2c0-4d18a354c711
ms.date: 06/08/2017
---


# Find.NoProofing Property (Word)

 **True** if Microsoft Word finds or replaces text that the spelling and grammar checker ignores. Read/write **Long** .


## Syntax

 _expression_ . **NoProofing**

 _expression_ A variable that represents a **[Find](find-object-word.md)** object.


## Example

This example searches for the string "hi" in text that the spelling and grammar checker ignores.


```vb
With Selection.Find 
 .ClearFormatting 
 .Text = "hi" 
 .NoProofing = True 
 .Execute Forward:=True 
End With
```


## See also


#### Concepts


[Find Object](find-object-word.md)

