---
title: Comment Object (PowerPoint)
keywords: vbapp10.chm642000
f1_keywords:
- vbapp10.chm642000
ms.prod: powerpoint
api_name:
- PowerPoint.Comment
ms.assetid: c1071b54-eeaa-0cec-13f0-b635da9511d8
ms.date: 06/08/2017
---


# Comment Object (PowerPoint)

Represents a comment on a given slide or slide range. The  **Comment** object is a member of the **[Comments](comments-object-powerpoint.md)** collection object.


## Remarks

Use the following properties to access comment data:


|||
|:-----|:-----|
|[Author](comment-author-property-powerpoint.md)|The author's full name|
|[AuthorIndex](comment-authorindex-property-powerpoint.md)|The author's index in the list of comments|
|[AuthorInitials](comment-authorinitials-property-powerpoint.md)|The author's initials|
|[DateTime](comment-datetime-property-powerpoint.md)|The date and time the comment was created|
|[Text](comment-text-property-powerpoint.md)|The text of the comment|
|[Left](comment-left-property-powerpoint.md), [Top](comment-top-property-powerpoint.md)|The comment's screen coordinates|

## Example

Use  **[Comments](slide-comments-property-powerpoint.md)** (index), where index is the number of the comment, or the **[Item](comments-item-method-powerpoint.md)** method to access a single comment on a slide. This example displays the author of the first comment on the first slide. If there are no comments, it displays a message stating such.


```vb
Sub ShowComment()

    With ActivePresentation.Slides(1).Comments

        If .Count > 0 Then

            MsgBox "The first comment on this slide is by " &; .Item(1).Author

        Else

            MsgBox "There are no comments on this slide."

        End If

    End With

End Sub
```

This example displays a message containing the author, date and time, and contents of all the messages on the first slide.




```vb
Sub SlideComments()

    Dim cmtExisting As Comment
    Dim cmtAll As Comments
    Dim strComments As String

    Set cmtAll = ActivePresentation.Slides(1).Comments

    If cmtAll.Count > 0 Then
        For Each cmtExisting In cmtAll
            strComments = strComments &; cmtExisting.Author &; vbTab &; _
                cmtExisting.DateTime &; vbTab &; cmtExisting.Text &; vbLf
        Next
        MsgBox "The comments in your document are as follows:" &; vbLf &; strComments
    Else
        MsgBox "This slide doesn't have any comments."
    End If

End Sub
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

