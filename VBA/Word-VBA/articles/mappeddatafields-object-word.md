---
title: MappedDataFields Object (Word)
keywords: vbawd10.chm2068
f1_keywords:
- vbawd10.chm2068
ms.prod: word
api_name:
- Word.MappedDataFields
ms.assetid: d67de1fb-f495-ff4a-f21d-fd165a96232c
ms.date: 06/08/2017
---


# MappedDataFields Object (Word)

A collection of  **MappedDataField** objects that represents all the mapped data fields available in Microsoft Word.


## Remarks

Use the  **MappedDataFields** property of the **MailMergeDataSource** object to return the **MappedDataFields** collection. This example creates a tabbed list of the mapped data fields available in Word and the fields in the data source to which they are mapped. This example assumes that the current document is a mail merge document and that the data source fields have corresponding mapped data fields.


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


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

