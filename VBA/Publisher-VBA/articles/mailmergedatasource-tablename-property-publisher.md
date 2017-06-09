---
title: MailMergeDataSource.TableName Property (Publisher)
keywords: vbapb10.chm6291491
f1_keywords:
- vbapb10.chm6291491
ms.prod: publisher
api_name:
- Publisher.MailMergeDataSource.TableName
ms.assetid: 0418bf66-550e-7dfc-671f-db2570a768d9
ms.date: 06/08/2017
---


# MailMergeDataSource.TableName Property (Publisher)

Returns a  **String** that represents the name of the table within the data source file that contains the mail merge records. The returned value may be blank if the table name is unknown or not applicable to the current data source. Read-only.


## Syntax

 _expression_. **TableName**

 _expression_A variable that represents a  **MailMergeDataSource** object.


### Return Value

String


## Example

This example displays a message with the name of the mail merge data source table name.


```vb
Sub EmployeeTable() 
 With ActiveDocument.MailMerge.DataSource 
 Select Case .TableName 
 Case "Employees" 
 MsgBox "This is an Employee mail merge publication." 
 Case "Customers" 
 MsgBox "This is a Customers mail merge publication." 
 Case "Suppliers" 
 MsgBox "This is a Suppliers mail merge publication." 
 Case "Shippers" 
 MsgBox "This is a Shippers mail merge publication." 
 Case Else 
 MsgBox "This is a " &; .TableName &; " mail merge publication." 
 End Select 
 End With 
End Sub
```


