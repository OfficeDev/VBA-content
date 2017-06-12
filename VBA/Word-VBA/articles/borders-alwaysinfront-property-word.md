---
title: Borders.AlwaysInFront Property (Word)
keywords: vbawd10.chm154927127
f1_keywords:
- vbawd10.chm154927127
ms.prod: word
api_name:
- Word.Borders.AlwaysInFront
ms.assetid: c005b911-47f6-fdc2-6098-4971b856b346
ms.date: 06/08/2017
---


# Borders.AlwaysInFront Property (Word)

 **True** if page borders are displayed in front of the document text. Read/write **Boolean** .


## Syntax

 _expression_ . **AlwaysInFront**

 _expression_ A variable that represents a **[Borders](borders-object-word.md)** object.


## Example

This example adds a graphical page border in front of text in the first section in the active document.


```vb
Dim borderLoop as Border 
 
With ActiveDocument.Sections(1) 
 .Borders.AlwaysInFront = True 
 For Each borderLoop In .Borders 
 With borderLoop 
 .ArtStyle = wdArtPeople 
 .ArtWidth = 15 
 End With 
 Next borderLoop 
End With
```


## See also


#### Concepts


[Borders Collection Object](borders-object-word.md)

