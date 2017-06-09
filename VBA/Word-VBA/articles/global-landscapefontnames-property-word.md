---
title: Global.LandscapeFontNames Property (Word)
keywords: vbawd10.chm163119116
f1_keywords:
- vbawd10.chm163119116
ms.prod: word
api_name:
- Word.Global.LandscapeFontNames
ms.assetid: 7c99f031-9290-1ff2-f2b6-da038a1c423b
ms.date: 06/08/2017
---


# Global.LandscapeFontNames Property (Word)

Returns a  **FontNames** object that includes the names of all the available landscape fonts.


## Syntax

 _expression_ . **LandscapeFontNames**

 _expression_ Required. A variable that represents a **[Global](global-object-word.md)** object.


## Example

This example creates a sorted list in a new document of the landscape font names in the  **FontNames** object.


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


[Global Object](global-object-word.md)

