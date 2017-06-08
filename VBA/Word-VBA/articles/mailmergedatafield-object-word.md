---
title: MailMergeDataField Object (Word)
keywords: vbawd10.chm2329
f1_keywords:
- vbawd10.chm2329
ms.prod: word
api_name:
- Word.MailMergeDataField
ms.assetid: ec0b8657-2842-73d2-5686-9f81b67a1871
ms.date: 06/08/2017
---


# MailMergeDataField Object (Word)

Represents a single mail merge field in a data source. The  **MailMergeDataField** object is a member of the **[MailMergeDataFields](mailmergedatafields-object-word.md)** collection. The **MailMergeDataFields** collection includes all the data fields in a mail merge data source (for example, Name, Address, and City).


## Remarks

Use  **DataFields** (Index), where Index is the data field name or the index number, to return a single **MailMergeDataField** object. The index number represents the position of the data field in the mail merge data source. The following example retrieves the first value from the FName field in the data source attached to the active document.


```
first = _ 
 ActiveDocument.MailMerge.DataSource.DataFields("FName").Value
```

The following example displays the name of first field in the data source attached to the active document.




```vb
MsgBox ActiveDocument.MailMerge.DataSource.DataFields(1).Name
```

You cannot add fields to the  **MailMergeDataFields** collection. All data fields in a data source are automatically included in the **MailMergeDataFields** collection.


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

