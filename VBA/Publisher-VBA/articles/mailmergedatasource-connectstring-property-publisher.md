---
title: MailMergeDataSource.ConnectString Property (Publisher)
keywords: vbapb10.chm6291460
f1_keywords:
- vbapb10.chm6291460
ms.prod: publisher
api_name:
- Publisher.MailMergeDataSource.ConnectString
ms.assetid: d7719567-f946-6b76-3ff2-d372dcc76a17
ms.date: 06/08/2017
---


# MailMergeDataSource.ConnectString Property (Publisher)

Returns a  **String** that represents the connection to the specified mail merge data source. Read-only.


## Syntax

 _expression_. **ConnectString**

 _expression_A variable that represents a  **MailMergeDataSource** object.


### Return Value

String


## Example

This example checks if the connection string contains the characters OLEDB and displays a message accordingly.


```vb
Sub VerifyCorrectDataSource() 
 
 With ActiveDocument.MailMerge.DataSource 
 If InStr(.ConnectString, "OLEDB") > 0 Then 
 MsgBox "OLE DB is used to connect to the data source." 
 Else 
 MsgBox "OLE DB is not used to connect to the data source." 
 End If 
 End With 
 
End Sub
```


