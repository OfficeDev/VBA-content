---
title: View.RevisionsBalloonWidth Property (Word)
keywords: vbawd10.chm161808424
f1_keywords:
- vbawd10.chm161808424
ms.prod: word
api_name:
- Word.View.RevisionsBalloonWidth
ms.assetid: f49d96e0-e159-38ec-fa61-6e5ba3827b1b
ms.date: 06/08/2017
---


# View.RevisionsBalloonWidth Property (Word)

Sets or returns a  **Single** representing the global setting in Microsoft Word that specifies the width of the revision balloons. Read/write.


## Syntax

 _expression_ . **RevisionsBalloonWidth**

 _expression_ An expression that returns one a **[View](view-object-word.md)** object.


## Remarks

The width of revision balloons includes padding of 0.5 inch between the document margin and the edge of the balloon and an eighth-inch between the edge of the balloon and the edge of the paper. Word adds space along the left or right edge of the paper. This width is extended into the margin and does not change the width of the document or paper size. Use the  **[RevisionsBalloonWidthType](view-revisionsballoonwidthtype-property-word.md)** property to specify the measurement to use when setting the **RevisionsBalloonWidth** property.


## Example

This example sets the width of the revision balloons to one inch and displays the revision balloons in the left margin. This example assumes that the document in the active window contains revisions made by one or more reviewers and that revisions are displayed in balloons.


```vb
Sub BalloonWidth() 
 With ActiveWindow.View 
 .RevisionsBalloonWidthType = wdBalloonWidthPoints 
 .RevisionsBalloonWidth = InchesToPoints(1) 
 .RevisionsBalloonSide = wdLeftMargin 
 End With 
End Sub
```


## See also


#### Concepts


[View Object](view-object-word.md)

