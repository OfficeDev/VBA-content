---
title: MailMerge.HighlightMergeFields Property (Word)
keywords: vbawd10.chm153092107
f1_keywords:
- vbawd10.chm153092107
ms.prod: word
api_name:
- Word.MailMerge.HighlightMergeFields
ms.assetid: 1002b34a-4492-97df-bb16-bd2c4319e055
ms.date: 06/08/2017
---


# MailMerge.HighlightMergeFields Property (Word)

 **True** to highlight the merge fields in a document. Read/write **Boolean** .


## Syntax

 _expression_ . **HighlightMergeFields**

 _expression_ A variable that represents a **[MailMerge](mailmerge-object-word.md)** object.


## Example

This example turns off highlighting merge fields in the active document.


```vb
Sub HighlightFields() 
 ActiveDocument.MailMerge.HighlightMergeFields = False 
End Sub
```


## See also


#### Concepts


[MailMerge Object](mailmerge-object-word.md)

