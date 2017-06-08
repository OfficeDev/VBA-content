---
title: Index.HeadingSeparator Property (Word)
keywords: vbawd10.chm159186945
f1_keywords:
- vbawd10.chm159186945
ms.prod: word
api_name:
- Word.Index.HeadingSeparator
ms.assetid: fa517204-b376-b25d-fbb2-8f1b5ef79e5c
ms.date: 06/08/2017
---


# Index.HeadingSeparator Property (Word)

Returns or sets the text between alphabetical groups (entries that start with the same letter) in the index. Corresponds to the \h switch for an INDEX field. Read/write  **WdHeadingSeparator** .


## Syntax

 _expression_ . **HeadingSeparator**

 _expression_ Required. A variable that represents an **[Index](index-object-word.md)** object.


## Example

This example formats the first index for the active document in a single column, with the appropriate letter preceding each alphabetical group.


```vb
If ActiveDocument.Indexes.Count >= 1 Then 
 With ActiveDocument.Indexes(1) 
 .HeadingSeparator = wdHeadingSeparatorLetter 
 .NumberOfColumns = 1 
 End With 
End If
```


## See also


#### Concepts


[Index Object](index-object-word.md)

