---
title: MailMergeDataSource.MappedDataFields Property (Word)
keywords: vbawd10.chm152895504
f1_keywords:
- vbawd10.chm152895504
ms.prod: word
api_name:
- Word.MailMergeDataSource.MappedDataFields
ms.assetid: 9fcebccf-976b-b57a-6cc3-c4d20fffea5c
ms.date: 06/08/2017
---


# MailMergeDataSource.MappedDataFields Property (Word)

Returns a  **[MappedDataFields](mappeddatafields-object-word.md)** collection that represents the mapped data fields available in Microsoft Word.


## Syntax

 _expression_ . **MappedDataFields**

 _expression_ An expression that returns a **[MailMergeDataSource](mailmergedatasource-object-word.md)** object.


## Example

This example creates a tabbed list of the mapped data fields available in Word and the fields in the data source to which they are mapped. This example assumes that the current document is a mail merge document.


```vb
Sub MappedFields() 
 Dim intCount As Integer 
 Dim docCurrent As Document 
 Dim docNew As Document 
 
 On Error Resume Next 
 
 Set docCurrent = ActiveDocument 
 Set docNew = Documents.Add 
 
 'Add leader tab to new document 
 docNew.Paragraphs.TabStops.Add _ 
 Position:=InchesToPoints(3.5), _ 
 Leader:=wdTabLeaderDots 
 
 With docCurrent.MailMerge.DataSource 
 
 'Insert heading paragraph for tabbed columns 
 docNew.Content.InsertAfter "Word Mapped Data Field" _ 
 &; vbTab &; "Data Source Field" 
 
 Do 
 intCount = intCount + 1 
 
 'Insert Word mapped data field name and the 
 'corresponding data source field name 
 docNew.Content.InsertAfter .MappedDataFields( _ 
 Index:=intCount).Name &; vbTab &; _ 
 .MappedDataFields(Index:=intCount) _ 
 .DataFieldName 
 
 'Insert paragraph 
 docNew.Content.InsertParagraphAfter 
 Loop Until intCount = .MappedDataFields.Count 
 
 End With 
 
End Sub
```


## See also


#### Concepts


[MailMergeDataSource Object](mailmergedatasource-object-word.md)

