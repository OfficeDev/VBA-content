---
title: Comment.DateTime Property (PowerPoint)
keywords: vbapp10.chm642006
f1_keywords:
- vbapp10.chm642006
ms.prod: powerpoint
api_name:
- PowerPoint.Comment.DateTime
ms.assetid: 52e08d04-18d6-61fc-1526-ef669aa5f6c8
ms.date: 06/08/2017
---


# Comment.DateTime Property (PowerPoint)

Returns the date and time a comment was created.


## Syntax

 _expression_. **DateTime**

 _expression_ A variable that represents a **Comment** object.


### Return Value

Date


## Remarks

Don't confuse this property with the  **[DateAndTime](headersfooters-dateandtime-property-powerpoint.md)** property, which applies to the headers and footers of a slide.


## Example

The following example provides information about all the comments for a given slide.


```vb
Sub ListComments()

    Dim cmtExisting As Comment
    Dim strAuthorInfo As String

    For Each cmtExisting In ActivePresentation.Slides(1).Comments
        With cmtExisting
            strAuthorInfo = strAuthorInfo &; .Author &; "'s comment #" &; _
                .AuthorIndex &; " (" &; .Text &; ") was created on " &; _
                .DateTime &; vbCrLf
        End With
    Next

    If strAuthorInfo <> "" Then
        MsgBox strAuthorInfo
    Else
        MsgBox "There are no comments on this slide."
    End If

End Sub
```


## See also


#### Concepts


[Comment Object](comment-object-powerpoint.md)

