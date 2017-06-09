---
title: Comment.AuthorIndex Property (PowerPoint)
keywords: vbapp10.chm642007
f1_keywords:
- vbapp10.chm642007
ms.prod: powerpoint
api_name:
- PowerPoint.Comment.AuthorIndex
ms.assetid: a004167b-a564-651e-1769-9e1a8947e385
ms.date: 06/08/2017
---


# Comment.AuthorIndex Property (PowerPoint)

Returns a  **Long** representing the index number of a comment for a given author. The first comment for a given author has an index number of 1, their second comment has an index number of 2. Read-only.


## Syntax

 _expression_. **AuthorIndex**

 _expression_ A variable that represents an **Comment** object.


### Return Value

Long


## Example

The following example provide information about the authors and their comment indexes for a given slide.


```vb
Sub GetCommentAuthorInfo()

    Dim cmtComment As Comment
    Dim strAuthorInfo As String

    With ActivePresentation.Slides(1)
        If .Comments.Count > 0 Then
            For Each cmtComment In .Comments
                strAuthorInfo = strAuthorInfo &; "Comment Number:  " &; _
                    cmtComment.AuthorIndex &; vbLf &; _
                    "Made by:  " &; cmtComment.Author &; vbLf &; _
                    "Says:  " &; cmtComment.Text &; vbLf &; vbLf
            Next cmtComment
        End If
    End With

    MsgBox "The comments for this slide are as follows: " &; _
        vbLf &; vbLf &; strAuthorInfo

End Sub
```


## See also


#### Concepts


[Comment Object](comment-object-powerpoint.md)

