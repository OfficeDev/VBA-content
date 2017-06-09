---
title: MailMerge.Check Method (Word)
keywords: vbawd10.chm153092202
f1_keywords:
- vbawd10.chm153092202
ms.prod: word
api_name:
- Word.MailMerge.Check
ms.assetid: a6f166e9-9c8c-80ec-9725-55efde2f4a3b
ms.date: 06/08/2017
---


# MailMerge.Check Method (Word)

Simulates the mail merge operation, pausing to report each error as it occurs.


## Syntax

 _expression_ . **Check**

 _expression_ Required. A variable that represents a **[MailMerge](mailmerge-object-word.md)** object.


## Example

This example checks the active document for mail merge errors.


```vb
Dim intState As Integer 
 
intState = ActiveDocument.MailMerge.State 
If intState = wdMainAndDataSource Or _ 
 intState = wdMainAndSourceAndHeader Then 
 ActiveDocument.MailMerge.Check 
End If
```


## See also


#### Concepts


[MailMerge Object](mailmerge-object-word.md)

