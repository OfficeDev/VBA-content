---
title: Borders.DistanceFrom Property (Word)
keywords: vbawd10.chm154927133
f1_keywords:
- vbawd10.chm154927133
ms.prod: word
api_name:
- Word.Borders.DistanceFrom
ms.assetid: 316858c5-51b0-1cc0-407f-0bee7d48aaae
ms.date: 06/08/2017
---


# Borders.DistanceFrom Property (Word)

Returns or sets a value that indicates whether the specified page border is measured from the edge of the page or from the text it surrounds. Read/write  **WdBorderDistanceFrom** .


## Syntax

 _expression_ . **DistanceFrom**

 _expression_ Required. A variable that represents a **[Borders](borders-object-word.md)** collection.


## Example

This example adds a single border around each page in section one in the active document and then sets the distance between each border and the corresponding edge of the page.


```vb
Dim borderLoop As Border 
 
With ActiveDocument.Sections(1) 
 For Each borderLoop In .Borders 
 borderLoop.LineStyle = wdLineStyleSingle 
 borderLoop.LineWidth = wdLineWidth050pt 
 Next borderLoop 
 With .Borders 
 .DistanceFrom = wdBorderDistanceFromPageEdge 
 .DistanceFromTop = 20 
 .DistanceFromLeft = 22 
 .DistanceFromBottom = 20 
 .DistanceFromRight = 22 
 End With 
End With
```

This example adds a border around each page in the first section in the selection, and then it sets the distance between the text and the page border to 6 points.




```vb
Dim borderLoop As Border 
 
With Selection.Sections(1) 
 For Each borderLoop In .Borders 
 borderLoop.ArtStyle = wdArtSeattle 
 borderLoop.ArtWidth = 22 
 Next borderLoop 
 With .Borders 
 .DistanceFrom = wdBorderDistanceFromText 
 .DistanceFromTop = 6 
 .DistanceFromLeft = 6 
 .DistanceFromBottom = 6 
 .DistanceFromRight = 6 
 End With 
End With
```


## See also


#### Concepts


[Borders Collection Object](borders-object-word.md)

