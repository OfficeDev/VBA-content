---
title: Borders.DistanceFromTop Property (Word)
keywords: vbawd10.chm154927108
f1_keywords:
- vbawd10.chm154927108
ms.prod: word
api_name:
- Word.Borders.DistanceFromTop
ms.assetid: 4e657225-0428-5d9f-582f-e2263fcd0437
ms.date: 06/08/2017
---


# Borders.DistanceFromTop Property (Word)

Returns or sets the space (in points) between the text and the top border. Read/write  **Long** .


## Syntax

 _expression_ . **DistanceFromTop**

 _expression_ A variable that represents a **[Borders](borders-object-word.md)** object.


## Remarks

Using this property with a page border, you can set either the space between the text and the top page border or the space between the top edge of the page and the top page border. Where the distance is measured from depends on the value of the  **[DistanceFrom](borders-distancefrom-property-word.md)** property.


## Example

This example adds a border around each paragraph in the selection and sets the distance between the text and the top border to 3 points.


```vb
With Selection.Borders 
 .Enable = True 
 .DistanceFromTop = 3 
End With
```

This example adds a border around each page in the first section in the selection. The example also sets the distance between the text and the page border to 6 points.




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

