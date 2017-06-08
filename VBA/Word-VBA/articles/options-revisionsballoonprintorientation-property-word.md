---
title: Options.RevisionsBalloonPrintOrientation Property (Word)
keywords: vbawd10.chm162988485
f1_keywords:
- vbawd10.chm162988485
ms.prod: word
api_name:
- Word.Options.RevisionsBalloonPrintOrientation
ms.assetid: ab97c3b8-2009-6171-9499-3d345f7b22e8
ms.date: 06/08/2017
---


# Options.RevisionsBalloonPrintOrientation Property (Word)

Returns or sets a  **WdRevisionsBalloonPrintOrientation** constant that represents the direction of revision and comment balloons when they are printed. Read/write.


## Syntax

 _expression_ . **RevisionsBalloonPrintOrientation**

 _expression_ Required. A variable that represents an **[Options](options-object-word.md)** collection.


## Example

This example prints documents with comments in Landscape format with the revision and comment balloons on one side of the page and the document text on the other.


```vb
Sub PrintLandscapeCommentBalloons() 
 Options.RevisionsBalloonPrintOrientation = _ 
 wdBalloonPrintOrientationForceLandscape 
End Sub
```


## See also


#### Concepts


[Options Object](options-object-word.md)

