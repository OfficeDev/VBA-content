---
title: Application.LandscapeFontNames Property (Word)
keywords: vbawd10.chm158334988
f1_keywords:
- vbawd10.chm158334988
ms.prod: word
api_name:
- Word.Application.LandscapeFontNames
ms.assetid: 59599ca0-0c6f-8d4a-9f4e-e98c5c241944
ms.date: 06/08/2017
---


# Application.LandscapeFontNames Property (Word)

Returns a  **[FontNames](fontnames-object-word.md)** object that includes the names of all the available landscape fonts.


## Syntax

 _expression_ . **LandscapeFontNames**

 _expression_ A variable that represents an **[Application](application-object-word.md)** object.


## Example

This example creates a sorted list in a new document of the landscape font names in the FontNames object.


```vb
Sub ListLandscapeFonts() 
 Dim docNew As Document 
 Dim intCount As Integer 
 
 Set docNew = Documents.Add 
 docNew.Content.InsertAfter "Landscape Fonts" &; vbLf 
 
 For intCount = 1 To LandscapeFontNames.Count 
 docNew.Content.InsertAfter LandscapeFontNames(intCount) _ 
 &; vbLf 
 Next 
 
 With docNew 
 .Range(Start:=.Paragraphs(2).Range.Start, End:=.Paragraphs _ 
 (docNew.Paragraphs.Count).Range.End).Select 
 End With 
 
 Selection.Sort 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-word.md)

