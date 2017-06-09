---
title: SlideRange.Comments Property (PowerPoint)
keywords: vbapp10.chm532032
f1_keywords:
- vbapp10.chm532032
ms.prod: powerpoint
api_name:
- PowerPoint.SlideRange.Comments
ms.assetid: ff06c024-66cf-d915-e0b0-676b009f93fb
ms.date: 06/08/2017
---


# SlideRange.Comments Property (PowerPoint)

Returns a  **[Comments](comments-object-powerpoint.md)** object that represents a collection of comments. Read-only.


## Syntax

 _expression_. **Comments**

 _expression_ A variable that represents a **SlideRange** object.


### Return Value

Comments


## Example

The following example adds a comment to a slide.


```vb
Sub AddNewComment()
    ActivePresentation.Slides(1).Comments.Add _
        Left:=0, Top:=0, Author:="John Doe", AuthorInitials:="jd", _
        Text:="Please check this spelling again before the next draft."
End Sub
```


## See also


#### Concepts


[SlideRange Object](sliderange-object-powerpoint.md)

