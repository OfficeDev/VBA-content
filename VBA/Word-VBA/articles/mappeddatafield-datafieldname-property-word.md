---
title: MappedDataField.DataFieldName Property (Word)
keywords: vbawd10.chm107544578
f1_keywords:
- vbawd10.chm107544578
ms.prod: word
api_name:
- Word.MappedDataField.DataFieldName
ms.assetid: 10356bc7-1635-8c83-984c-72a332740d89
ms.date: 06/08/2017
---


# MappedDataField.DataFieldName Property (Word)

Sets or returns a  **String** that represents the name of the field in the mail merge data source to which a mapped data field maps. Read/write.


## Syntax

 _expression_ . **DataFieldName**

 _expression_ An expression that returns a **[MappedDataField](mappeddatafield-object-word.md)** object.


## Remarks

A blank string is returned if the specified data field is not mapped to a mapped data field.


## Example

This example creates a tabbed list of the mapped data fields available in Word and the fields in the data source to which they are mapped. This example assumes that the current document is a mail merge document and that the data source fields have corresponding mapped data fields.


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


[MappedDataField Object](mappeddatafield-object-word.md)

