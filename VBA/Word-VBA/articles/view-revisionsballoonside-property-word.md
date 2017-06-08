---
title: View.RevisionsBalloonSide Property (Word)
keywords: vbawd10.chm161808426
f1_keywords:
- vbawd10.chm161808426
ms.prod: word
api_name:
- Word.View.RevisionsBalloonSide
ms.assetid: 629d67a3-49c3-82f0-01af-c93913f9e268
ms.date: 06/08/2017
---


# View.RevisionsBalloonSide Property (Word)

Sets or returns a  **WdRevisionsBalloonMargin** constant that specifies whether Word displays revision balloons in the left or right margin in a document.


## Syntax

 _expression_ . **RevisionsBalloonSide**

 _expression_ Required. A variable that represents a **[View](view-object-word.md)** object.


## Example

This example switches the revision balloons between the left side and the right side. This example assumes that the document in the active window contains revisions made by one or more reviewers and that revisions are displayed in balloons.


```vb
Sub ToggleRevisionBalloons() 
 With ActiveWindow.View 
 If .RevisionsBalloonSide = wdLeftMargin Then 
 .RevisionsBalloonSide = wdRightMargin 
 Else 
 .RevisionsBalloonSide = wdLeftMargin 
 End If 
 End With 
End Sub
```


## See also


#### Concepts


[View Object](view-object-word.md)

