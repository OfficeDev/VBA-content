---
title: MailMergeField.Cut Method (Word)
keywords: vbawd10.chm152961130
f1_keywords:
- vbawd10.chm152961130
ms.prod: word
api_name:
- Word.MailMergeField.Cut
ms.assetid: 83455a23-06cb-9c73-1655-ad6c86d9cb3b
ms.date: 06/08/2017
---


# MailMergeField.Cut Method (Word)

Removes the specified mail merge field from the document and moves it to the Clipboard.


## Syntax

 _expression_ . **Cut**

 _expression_ Required. A variable that represents a **[MailMergeField](mailmergefield-object-word.md)** object.


## Example

This example deletes the first field in the active document and pastes the field at the insertion point.


```vb
If ActiveDocument.Fields.Count >= 1 Then 
 ActiveDocument.Fields(1).Cut 
 Selection.Collapse Direction:=wdCollapseEnd 
 Selection.Paste 
End If
```


## See also


#### Concepts


[MailMergeField Object](mailmergefield-object-word.md)

