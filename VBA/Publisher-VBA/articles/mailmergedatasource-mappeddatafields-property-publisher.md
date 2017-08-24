---
title: MailMergeDataSource.MappedDataFields Property (Publisher)
keywords: vbapb10.chm6291475
f1_keywords:
- vbapb10.chm6291475
ms.prod: publisher
api_name:
- Publisher.MailMergeDataSource.MappedDataFields
ms.assetid: 9f2a15a7-41b0-6025-73d6-eb70a412b830
ms.date: 06/08/2017
---


# MailMergeDataSource.MappedDataFields Property (Publisher)

Returns a  **[MailMergeMappedDataFields](mailmergemappeddatafields-object-publisher.md)** object that represents the mapped data fields available in Microsoft Publisher.


## Syntax

 _expression_. **MappedDataFields**

 _expression_A variable that represents a  **MailMergeDataSource** object.


### Return Value

MailMergeMappedDataFields


## Example

This example creates a table on a new page of the current publication and lists the mapped data fields available in Publisher and the fields in the data source to which they are mapped. This example assumes that the current publication is a mail merge publication and that the data source fields have corresponding mapped data fields.


```vb
Sub MappedFields() 
 Dim intCount As Integer 
 Dim intRows As Integer 
 Dim docPub As Document 
 Dim pagNew As Page 
 Dim shpTable As Shape 
 Dim tblTable As Table 
 Dim rowTable As Row 
 
 On Error Resume Next 
 
 Set docPub = ThisDocument 
 Set pagNew = ThisDocument.Pages.Add(Count:=1, After:=1) 
 intRows = docPub.MailMerge.DataSource.MappedDataFields.Count + 1 
 
 'Creates new table with a heading row 
 Set shpTable = pagNew.Shapes.AddTable(NumRows:=intRows, _ 
 numColumns:=2, Left:=100, Top:=100, Width:=400, Height:=12) 
 Set tblTable = shpTable.Table 
 With tblTable.Rows(1) 
 With .Cells(1).Text 
 .Text = "Mapped Data Field" 
 .Font.Bold = msoTrue 
 End With 
 With .Cells(2).Text 
 .Text = "Data Source Field" 
 .Font.Bold = msoTrue 
 End With 
 End With 
 
 With docPub.MailMerge.DataSource 
 For intCount = 2 To intRows - 1 
 'Inserts mapped data field name and the 
 'corresponding data source field name 
 tblTable.Rows(intCount - 1).Cells(1).Text _ 
 .Text = .MappedDataFields(Index:=intCount).Name 
 tblTable.Rows(intCount - 1).Cells(2).Text _ 
 .Text = .MappedDataFields(Index:=intCount).DataFieldName 
 Next 
 End With 
End Sub
```


