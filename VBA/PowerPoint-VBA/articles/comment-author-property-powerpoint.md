---
title: Comment.Author Property (PowerPoint)
keywords: vbapp10.chm642003
f1_keywords:
- vbapp10.chm642003
ms.prod: powerpoint
api_name:
- PowerPoint.Comment.Author
ms.assetid: 83feff12-02a1-444e-baaf-15e39049e6a5
ms.date: 06/08/2017
---


# Comment.Author Property (PowerPoint)

Returns a  **String** that represents the author as for a specified **[Comment](comment-object-powerpoint.md)** object. Read-only.


## Syntax

 _expression_. **Author**

 _expression_ A variable that represents an **Comment** object.


### Return Value

String


## Remarks

This property returns only the author's name. To return the author's initials, use the  **[AuthorInitials](comment-authorinitials-property-powerpoint.md)** property. Specify the author of a comment when you add a new comment to the presentation.


## Example

The following example adds a comment to the first slide of the active presentation and then displays the author's name and initials in a message.


```vb
Sub GetAuthorName()

    With ActivePresentation.Slides(1)
        .Comments.Add Left:=100, Top:=100, Author:="Jeff Smith", _
            AuthorInitials:="JS", _
            Text:="This is a new comment added to the first slide."
        MsgBox "This comment was created by " &; _
            .Comments(1).Author &; " (" &; .Comments(1).AuthorInitials &; ")."
    End With
	
End Sub
```


## See also


#### Concepts


[Comment Object](comment-object-powerpoint.md)

