---
title: MailMergeFieldName Object (Word)
keywords: vbawd10.chm2331
f1_keywords:
- vbawd10.chm2331
ms.prod: word
api_name:
- Word.MailMergeFieldName
ms.assetid: f4e09d1e-0da2-2f0f-1747-566a4ae443b6
ms.date: 06/08/2017
---


# MailMergeFieldName Object (Word)

Represents a mail merge field name in a data source. The  **MailMergeFieldName** object is a member of the **[MailMergeFieldNames](mailmergefieldnames-object-word.md)** collection. The **MailMergeFieldNames** collection includes all the data field names in a mail merge data source.


## Remarks

Use  **FieldNames** (Index), where Index is the index number, to return a single **MailMergeFieldName** object. The index number represents the position of the field in the mail merge data source. The following example retrieves the name of the last field in the data source attached to the active document.


```
alast = ActiveDocument.MailMerge.DataSource.FieldNames.Count 
afirst = ActiveDocument.MailMerge.DataSource.FieldNames(alast).Name 
MsgBox afirst
```

You cannot add fields to the  **MailMergeFieldNames** collection. Field names in a data source are automatically included in the **MailMergeFieldNames** collection.


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

