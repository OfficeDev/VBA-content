---
title: View.RevisionsBalloonWidthType Property (Word)
keywords: vbawd10.chm161808425
f1_keywords:
- vbawd10.chm161808425
ms.prod: word
api_name:
- Word.View.RevisionsBalloonWidthType
ms.assetid: f300fc90-df18-cef4-bc00-dce76f2feff8
ms.date: 06/08/2017
---


# View.RevisionsBalloonWidthType Property (Word)

Sets or returns a  **WdRevisionsBalloonWidthType** constant representing the global setting that specifies how Microsoft Word measures the width of revision balloons. Read/write.


## Syntax

 _expression_ . **RevisionsBalloonWidthType**

 _expression_ Required. A variable that represents a **[View](view-object-word.md)** object.


## Remarks

The  **RevisionsBalloonWidthType** property sets the measurement unit to use when setting the **RevisionsBalloonWidth** property.


## Example

This example sets the width of the revision balloons to twenty-five percent of the document's width. This example assumes that the document in the active window contains revisions made by one or more reviewers and that revisions are displayed in balloons.


```vb
Sub BalloonWidthType() 
 With ActiveWindow.View 
 .RevisionsBalloonWidthType = wdBalloonWidthPercent 
 .RevisionsBalloonWidth = 25 
 End With 
End Sub
```


## See also


#### Concepts


[View Object](view-object-word.md)

