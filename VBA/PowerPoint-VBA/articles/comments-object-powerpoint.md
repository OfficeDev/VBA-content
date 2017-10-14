---
title: Comments Object (PowerPoint)
keywords: vbapp10.chm641000
f1_keywords:
- vbapp10.chm641000
ms.prod: powerpoint
api_name:
- PowerPoint.Comments
ms.assetid: 1f29db7c-90fa-db9f-5229-136534ce803d
ms.date: 06/08/2017
---


# Comments Object (PowerPoint)

Represents a collection of  **[Comment](comment-object-powerpoint.md)** objects.


## Example

Use the [Comments](slide-comments-property-powerpoint.md)property to refer to the  **Comments** collection. The following example displays the number of comments on the current slide.


```vb
Sub CountComments()
    MsgBox "You have " &; ActiveWindow.Selection.SlideRange(1) _
        .Comments.Count &; " comments on this slide."
End Sub
```

Use the [Add](comments-add-method-powerpoint.md)method to add a comment to a slide. This example adds a new comment to the first slide of the active presentation.




```vb
Sub AddComment()

    Dim sldNew As Slide
    Dim cmtNew As Comment

    Set sldNew = ActivePresentation.Slides.Add(Index:=1, _
        Layout:=ppLayoutBlank)

    Set cmtNew = sldNew.Comments.Add(Left:=12, Top:=12, _
        Author:="Jeff Smith", AuthorInitials:="JS", _
        Text:="You might consider reviewing the new specs" &; _
        "for more up-to-date information.")

End Sub
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

