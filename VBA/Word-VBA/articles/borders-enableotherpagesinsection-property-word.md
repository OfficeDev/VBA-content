---
title: Borders.EnableOtherPagesInSection Property (Word)
keywords: vbawd10.chm154927135
f1_keywords:
- vbawd10.chm154927135
ms.prod: word
api_name:
- Word.Borders.EnableOtherPagesInSection
ms.assetid: 288caacd-e8c8-fa27-fda0-0d02932b90aa
ms.date: 06/08/2017
---


# Borders.EnableOtherPagesInSection Property (Word)

 **True** if page borders are enabled for all pages in the section except for the first page. Read/write **Boolean** .


## Syntax

 _expression_ . **EnableOtherPagesInSection**

 _expression_ A variable that represents a **[Borders](borders-object-word.md)** object.


## Example

This example adds a border around each page in the first section in the selection except for the first page.


```vb
Dim borderLoop As Border 
 
With Selection.Sections(1) 
 .Borders.EnableFirstPageInSection = False 
 .Borders.EnableOtherPagesInSection = True 
 For Each borderLoop In .Borders 
 borderLoop.ArtStyle = wdArtBabyRattle 
 borderLoop.ArtWidth = 22 
 Next borderLoop 
End With
```


## See also


#### Concepts


[Borders Collection Object](borders-object-word.md)

