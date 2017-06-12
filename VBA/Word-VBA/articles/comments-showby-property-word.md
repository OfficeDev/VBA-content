---
title: Comments.ShowBy Property (Word)
keywords: vbawd10.chm155190251
f1_keywords:
- vbawd10.chm155190251
ms.prod: word
api_name:
- Word.Comments.ShowBy
ms.assetid: 13568867-ca6b-828a-1914-f6f32099b976
ms.date: 06/08/2017
---


# Comments.ShowBy Property (Word)

Returns or sets the name of the reviewer whose comments are shown in the comments pane. Read/write  **String** .


## Syntax

 _expression_ . **ShowBy**

 _expression_ An expression that returns a **[Comments](comments-object-word.md)** collection object.


## Remarks

You can choose to show comments either by a single reviewer or by all reviewers. To view the comments by all reviewers, set this property to "All Reviewers."


## Example

The following example displays comments made by Don Funk.


```vb
If ActiveDocument.Comments.Count >= 1 Then 
 ActiveDocument.ActiveWindow.View.SplitSpecial = wdPaneComments 
 ActiveDocument.Comments.ShowBy = "Don Funk" 
End If
```


## See also


#### Concepts


[Comments Collection Object](comments-object-word.md)

