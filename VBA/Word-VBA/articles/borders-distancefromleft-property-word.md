---
title: Borders.DistanceFromLeft Property (Word)
keywords: vbawd10.chm154927124
f1_keywords:
- vbawd10.chm154927124
ms.prod: word
api_name:
- Word.Borders.DistanceFromLeft
ms.assetid: 614f44d6-3214-ad4b-42e5-f42c09f180f4
ms.date: 06/08/2017
---


# Borders.DistanceFromLeft Property (Word)

Returns or sets the space (in points) between the text and the left border. Read/write  **Long** .


## Syntax

 _expression_ . **DistanceFromLeft**

 _expression_ A variable that represents a **[Borders](borders-object-word.md)** object.


## Remarks

Using this property with a page border, you can set either the space between the text and the left page border or the space between the left edge of the page and the left page border. Where the distance is measured from depends on the value of the  **[DistanceFrom](borders-distancefrom-property-word.md)** property.


## Example

This example adds a border around each frame in the active document and sets the distance between the frame and the border to 5 points.


```vb
Dim frameLoop As Frame 
 
For Each frameLoop In ActiveDocument.Frames 
 With frameLoop.Borders 
 .Enable = True 
 .DistanceFromLeft = 5 
 .DistanceFromRight = 5 
 .DistanceFromTop = 5 
 .DistanceFromBottom = 5 
 End With 
Next frameLoop
```

This example adds a border around the first paragraph in the active document and sets the distance between the text and the left border to 3 points.




```vb
With ActiveDocument.Paragraphs(1).Borders 
 .Enable = True 
 .DistanceFromLeft = 3 
End With
```


## See also


#### Concepts


[Borders Collection Object](borders-object-word.md)

