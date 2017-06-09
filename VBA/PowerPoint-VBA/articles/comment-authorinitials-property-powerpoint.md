---
title: Comment.AuthorInitials Property (PowerPoint)
keywords: vbapp10.chm642004
f1_keywords:
- vbapp10.chm642004
ms.prod: powerpoint
api_name:
- PowerPoint.Comment.AuthorInitials
ms.assetid: 21789206-78b0-2c9e-4461-5005d821bd6c
ms.date: 06/08/2017
---


# Comment.AuthorInitials Property (PowerPoint)

Returns the author's initials as a read-only  **String** for a specified **[Comment](comment-object-powerpoint.md)** object. Read-only.


## Syntax

 _expression_. **AuthorInitials**

 _expression_ A variable that represents an **Comment** object.


### Return Value

String


## Remarks

This property only returns the author's initials. To return the author's name use the  **[Author](comment-author-property-powerpoint.md)** property. Specify the author's initials when you add a new comment to the presentation.


## Example

The following example returns the author's initials for a specified comment.


```vb
Sub GetAuthorName()

    With ActivePresentation.Slides(1)
        .Comments.Add Left:=100, Top:=100, Author:="Jeff Smith", _
            AuthorInitials:="JS", _
            Text:="This is a new comment added to the first slide."

        MsgBox .Comments(1).Author &; .Comments(1).AuthorInitials
    End With

End Sub
```


## See also


#### Concepts


[Comment Object](comment-object-powerpoint.md)

