---
title: ListLevel.ApplyPictureBullet Method (Word)
keywords: vbawd10.chm160235520
f1_keywords:
- vbawd10.chm160235520
ms.prod: word
api_name:
- Word.ListLevel.ApplyPictureBullet
ms.assetid: 9d91b047-c91b-60e1-2b57-aaa16491d212
ms.date: 06/08/2017
---


# ListLevel.ApplyPictureBullet Method (Word)

Formats a paragraph or range of paragraphs with a picture bullet.


## Syntax

 _expression_ . **ApplyPictureBullet**( **_FileName_** )

 _expression_ Required. A variable that represents a **[ListLevel](listlevel-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The path and file name of the picture file.|

## Example

This example creates a new document with a list and applies a picture bullet format to all paragraphs except the first and last.


```vb
Sub ApplyPictureBulletsToParagraphs() 
 Dim docNew As Document 
 Dim rngRange As Range 
 Dim lstTemplate As ListTemplate 
 Dim intPara As Integer 
 Dim intCount As Integer 
 
 'Set the initial value of object variables 
 Set docNew = Documents.Add 
 
 'Add paragraphs to the new document, including eight list items 
 With Selection 
 .TypeText Text:="This is an introductory paragraph." 
 .TypeParagraph 
 End With 
 Do Until intPara = 8 
 With Selection 
 .TypeText Text:="This is a list item." 
 .TypeParagraph 
 End With 
 intPara = intPara + 1 
 Loop 
 Selection.TypeText Text:="This is a concluding paragraph." 
 
 'Count the total number of paragraphs in the document 
 intCount = docNew.Paragraphs.Count 
 
 'Set the range to include all paragraphs in the 
 'document except the first and the last 
 Set rngRange = docNew.Range( _ 
 Start:=ActiveDocument.Paragraphs(2).Range.Start, _ 
 End:=ActiveDocument.Paragraphs(intCount - 1).Range.End) 
 
 'Format the list template as a bullet 
 Set lstTemplate = ListGalleries(Index:=wdBulletGallery) _ 
 .ListTemplates(7) 
 
 'Format list with a picture bullet 
 lstTemplate.ListLevels(1).ApplyPictureBullet _ 
 FileName:="c:\pictures\bluebullet.gif" 
 
 'Apply the list format settings to the range 
 rngRange.ListFormat.ApplyListTemplate _ 
 ListTemplate:=lstTemplate 
End Sub
```


## See also


#### Concepts


[ListLevel Object](listlevel-object-word.md)

