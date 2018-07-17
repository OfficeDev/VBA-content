---
title: Document.IsDataSourceConnected Property (Publisher)
keywords: vbapb10.chm196722
f1_keywords:
- vbapb10.chm196722
ms.prod: publisher
api_name:
- Publisher.Document.IsDataSourceConnected
ms.assetid: b62422ab-12f7-1151-d8d1-1cb32de18160
ms.date: 06/08/2017
---


# Document.IsDataSourceConnected Property (Publisher)

 **True** if the specified publication is connected to a data source. Read-only.


## Syntax

 _expression_. **IsDataSourceConnected**

 _expression_A variable that represents an  **Document** object.


## Remarks

A publication must be connected to a valid data source to perform a mail merge or catalog merge.


## Example

The following example tests whether the publication is connected to a data source and, if it is not, specifies and connects a data source to the publication. 

Before running this example, you must replace  _PathToFile_ with a valid file path and _TableName_ with a valid data source table name.




```vb
Dim strDataSource As String 
Dim strDataSourceTable As String 
 
 'Specify data source and table name 
 
 strDataSource = "PathToFile" 
 strDataSourceTable = "TableName" 
 
 'Connect to a datasource 
 If Not (ThisDocument.IsDataSourceConnected) Then 
 ThisDocument.MailMerge.OpenDataSource strDataSource, , strDataSourceTable 
 
 End If
```


