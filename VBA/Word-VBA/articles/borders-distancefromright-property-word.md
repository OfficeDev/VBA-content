---
title: Borders.DistanceFromRight Property (Word)
keywords: vbawd10.chm154927126
f1_keywords:
- vbawd10.chm154927126
ms.prod: word
api_name:
- Word.Borders.DistanceFromRight
ms.assetid: 456510ef-6746-6ef2-68a9-6917ce59144d
ms.date: 06/08/2017
---


# Borders.DistanceFromRight Property (Word)

Returns or sets the space (in points) between the right edge of the text and the right border. Read/write  **Long** .


## Syntax

 _expression_ . **DistanceFromRight**

 _expression_ A variable that represents a **[Borders](borders-object-word.md)** object.


## Remarks

Using this property with a page border, you can set either the space between the text and the right border or the space between the right edge of the page and the right border. Where the distance is measured from depends on the value of the  **[DistanceFrom](borders-distancefrom-property-word.md)** property.


## Example

This example adds a border around each paragraph in the selection and sets the distance between the text and the right border to 3 points.


```vb
With Selection.Paragraphs.Borders 
 .Enable = True 
 .DistanceFromRight = 3 
End With
```

This example adds a single border around each page in section one in the active document. The example also sets the distance between the right and left border and the corresponding edges of the page to 22 points.




```vb
Dim borderLoop As Border 
 
With ActiveDocument.Sections(1) 
 For Each borderLoop In .Borders 
 borderLoop.LineStyle = wdLineStyleSingle 
 borderLoop.LineWidth = wdLineWidth050pt 
 Next borderLoop 
 With .Borders 
 .DistanceFrom = wdBorderDistanceFromPageEdge 
 .DistanceFromLeft = 22 
 .DistanceFromRight = 22 
 End With 
End With
```


## See also


#### Concepts


[Borders Collection Object](borders-object-word.md)

