---
title: Slide.Comments Property (PowerPoint)
keywords: vbapp10.chm531028
f1_keywords:
- vbapp10.chm531028
ms.prod: powerpoint
api_name:
- PowerPoint.Slide.Comments
ms.assetid: 396c2d6b-f0cb-3ed8-94ae-6ee864d194c1
ms.date: 06/08/2017
---


# Slide.Comments Property (PowerPoint)

Returns a  **[Comments](comments-object-powerpoint.md)** object that represents a collection of comments. Read-only.


## Syntax

 _expression_. **Comments**

 _expression_ A variable that represents a **Slide** object.


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


[Slide Object](slide-object-powerpoint.md)

