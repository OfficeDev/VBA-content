---
title: MailMergeDataSource.ConnectString Property (Word)
keywords: vbawd10.chm152895493
f1_keywords:
- vbawd10.chm152895493
ms.prod: word
api_name:
- Word.MailMergeDataSource.ConnectString
ms.assetid: e402bc58-89e8-f18a-f70d-d970297777be
ms.date: 06/08/2017
---


# MailMergeDataSource.ConnectString Property (Word)

Returns the connection string for the specified mail merge data source. Read-only  **String** .


## Syntax

 _expression_ . **ConnectString**

 _expression_ A variable that represents a **[MailMergeDataSource](mailmergedatasource-object-word.md)** object.


## Example

This example creates a new main document and attaches the Customers table from a Microsoft Access database named "Northwind.mdb." The connection string is displayed in a message box.


```vb
Dim docNew As Document 
 
Set docNew = Documents.Add 
 
With docNew.MailMerge 
 .MainDocumentType = wdFormLetters 
 .OpenDataSource _ 
 Name:="C:\Program Files\Microsoft Office\Office" &; _ 
 "\Samples\Northwind.mdb", _ 
 LinkToSource:=True, AddToRecentFiles:=False, _ 
 Connection:="TABLE Customers" 
 MsgBox .DataSource.ConnectString 
End With
```


## See also


#### Concepts


[MailMergeDataSource Object](mailmergedatasource-object-word.md)

