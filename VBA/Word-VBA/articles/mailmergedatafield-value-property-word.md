---
title: MailMergeDataField.Value Property (Word)
keywords: vbawd10.chm152633344
f1_keywords:
- vbawd10.chm152633344
ms.prod: word
api_name:
- Word.MailMergeDataField.Value
ms.assetid: 742c8cea-3313-67d1-2f62-b4730cd753ab
ms.date: 06/08/2017
---


# MailMergeDataField.Value Property (Word)

Returns the contents of the mail merge data field or mapped data field for the current record. Read-only  **String** .


## Syntax

 _expression_ . **Value**

 _expression_ Required. A variable that represents a **[MailMergeDataField](mailmergedatafield-object-word.md)** object.


## Remarks

Use the  **[ActiveRecord](mailmergedatasource-activerecord-property-word.md)** property to set the active record in a mail merge data source.


## Example

This example displays the contents of the active record in the data source attached to Main.doc.


```vb
For Each dataF In _ 
 Documents("Main.doc").MailMerge.DataSource.DataFields 
 If dataF.Value <> "" Then dRecord = dRecord &; _ 
 dataF.Value &; vbCr 
Next dataF 
MsgBox dRecord
```


## See also


#### Concepts


[MailMergeDataField Object](mailmergedatafield-object-word.md)

