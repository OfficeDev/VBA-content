---
title: Border.ArtStyle Property (Word)
keywords: vbawd10.chm154861573
f1_keywords:
- vbawd10.chm154861573
ms.prod: word
api_name:
- Word.Border.ArtStyle
ms.assetid: 999569c0-96de-0c6c-462c-ec32804f8801
ms.date: 06/08/2017
---


# Border.ArtStyle Property (Word)

Returns or sets the graphical page-border design for a document. Read/write  **WdPageBorderArt** .


## Syntax

 _expression_ . **ArtStyle**

 _expression_ Required. A variable that represents a **[Border](border-object-word.md)** object.


## Example

This example adds a border of black dots around each page in first section in the selection.


```vb
Dim borderLoop As Border 
 
For Each borderLoop In Selection.Sections(1).Borders 
 With borderLoop 
 .ArtStyle = wdArtBasicBlackDots 
 .ArtWidth = 6 
 End With 
Next borderLoop
```

This example adds a picture border around each page in section one in the active document.




```vb
Dim borderLoop As Border 
 
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


[Border Object](border-object-word.md)

