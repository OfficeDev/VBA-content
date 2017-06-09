---
title: MailMerge.DataSource Property (Word)
keywords: vbawd10.chm153092100
f1_keywords:
- vbawd10.chm153092100
ms.prod: word
api_name:
- Word.MailMerge.DataSource
ms.assetid: d05103ce-3d5a-74e5-d21a-d58eb5bbf992
ms.date: 06/08/2017
---


# MailMerge.DataSource Property (Word)

Returns a  **[MailMergeDataSource](mailmergedatasource-object-word.md)** object that refers to the data source attached to a mail merge main document. Read-only.


## Syntax

 _expression_ . **DataSource**

 _expression_ A variable that represents a **[MailMerge](mailmerge-object-word.md)** object.


## Example

This example displays the name of the data source attached to the active document.


```vb
If ActiveDocument.MailMerge.DataSource.Name <> "" Then _ 
 MsgBox ActiveDocument.MailMerge.DataSource.Name
```

This example displays the next record from the data source attached to Main.doc.




```vb
ActiveDocument.ActiveWindow.View.ShowFieldCodes = False 
With Documents("Main.doc").MailMerge 
 .ViewMailMergeFieldCodes = False 
 .DataSource.ActiveRecord = wdNextRecord 
End With
```


## See also


#### Concepts


[MailMerge Object](mailmerge-object-word.md)

