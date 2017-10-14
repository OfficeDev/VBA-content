---
title: MailMergeDataSource.OpenRecipientsDialog Method (Publisher)
keywords: vbapb10.chm6291490
f1_keywords:
- vbapb10.chm6291490
ms.prod: publisher
api_name:
- Publisher.MailMergeDataSource.OpenRecipientsDialog
ms.assetid: 5a0a2b4a-ce23-435c-6e18-f778d6e14fd6
ms.date: 06/08/2017
---


# MailMergeDataSource.OpenRecipientsDialog Method (Publisher)

Displays the  **Recipients** dialog box for a mail merge publication.


## Syntax

 _expression_. **OpenRecipientsDialog**

 _expression_A variable that represents a  **MailMergeDataSource** object.


## Example

This example displays the  **Mail Merge Recipients** dialog box.


```vb
Sub ShowRecipientsDialog() 
 ActiveDocument.MailMerge.DataSource.OpenRecipientsDialog 
End Sub
```


