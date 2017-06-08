---
title: Hyperlinks Object (PowerPoint)
keywords: vbapp10.chm525000
f1_keywords:
- vbapp10.chm525000
ms.prod: powerpoint
api_name:
- PowerPoint.Hyperlinks
ms.assetid: 33a3fe49-6302-0f53-22f6-b8b1594d5d57
ms.date: 06/08/2017
---


# Hyperlinks Object (PowerPoint)

A collection of all the  **[Hyperlink](hyperlink-object-powerpoint.md)** objects on a slide or master.


## Example

Use the [Hyperlinks](slide-hyperlinks-property-powerpoint.md)property to return the  **Hyperlinks** collection. The following example updates all hyperlinks on slide one in the active presentation that have the specified address.


```vb
For Each hl In ActivePresentation.Slides(1).Hyperlinks

    If hl.Address = "c:\current work\sales.ppt" Then

        hl.Address = "c:\new\newsales.ppt"

    End If

Next
```

Use the [Hyperlink](actionsetting-hyperlink-property-powerpoint.md)property to create a hyperlink and add it to the  **Hyperlinks** collection. The following example sets a hyperlink that will be followed when the user clicks shape three on slide one in the active presentation during a slide show and adds the new hyperlink to the collection. Note that if shape three already has a mouse-click hyperlink defined, the following example will delete this hyperlink from the collection when it adds the new one, so the number of items in the **Hyperlinks** collection won't change.




```vb
With ActivePresentation.Slides(1).Shapes(3) _
        .ActionSettings(ppMouseClick)

    .Action = ppActionHyperlink
    .Hyperlink.Address = "http://www.microsoft.com"

End With
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

