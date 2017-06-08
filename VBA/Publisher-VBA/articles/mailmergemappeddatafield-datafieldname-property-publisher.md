---
title: MailMergeMappedDataField.DataFieldName Property (Publisher)
keywords: vbapb10.chm6553603
f1_keywords:
- vbapb10.chm6553603
ms.prod: publisher
api_name:
- Publisher.MailMergeMappedDataField.DataFieldName
ms.assetid: c30e56c1-c4f4-a581-00d1-eb367178e0af
ms.date: 06/08/2017
---


# MailMergeMappedDataField.DataFieldName Property (Publisher)

Returns or sets a  **String** which represents the name of the field in the mail merge data source to which a mapped data field maps. An empty string is returned if the specified data field is not mapped to a mapped data field. Read/write.


## Syntax

 _expression_. **DataFieldName**

 _expression_A variable that represents a  **MailMergeMappedDataField** object.


### Return Value

String


## Example

This example creates a table on a new page of the current publication and lists the mapped data fields available and the fields in the data source to which they are mapped. This example assumes that the current publication is a mail merge publication and that the data source fields have corresponding mapped data fields.


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


