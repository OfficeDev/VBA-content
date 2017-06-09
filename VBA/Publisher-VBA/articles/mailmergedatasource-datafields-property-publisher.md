---
title: MailMergeDataSource.DataFields Property (Publisher)
keywords: vbapb10.chm6291461
f1_keywords:
- vbapb10.chm6291461
ms.prod: publisher
api_name:
- Publisher.MailMergeDataSource.DataFields
ms.assetid: 820af882-d54c-a205-2925-e7110fc0c02b
ms.date: 06/08/2017
---


# MailMergeDataSource.DataFields Property (Publisher)

Returns a  **[MailMergeDataFields](mailmergedatafields-object-publisher.md)** collection that represents the fields in the specified data source.


## Syntax

 _expression_. **DataFields**

 _expression_A variable that represents a  **MailMergeDataSource** object.


### Return Value

MailMergeDataFields


## Example

This example displays the value of the value of the FirstName and LastName fields from the active record in the data source attached to the active publication.


```vb
Sub ShowNameForActiveRecord() 
 Dim mdfFirst As MailMergeDataField 
 Dim mdfLast As MailMergeDataField 
 
 With ActiveDocument.MailMerge.DataSource 
 Set mdfFirst = .DataFields.Item("FirstName") 
 Set mdfLast = .DataFields.Item("LastName") 
 MsgBox "The active record in the attached " &; _ 
 vbLf &; "data source is : " &; _ 
 mdfFirst.Value &; " " &; _ 
 mdfLast.Value 
 End With 
End Sub
```


