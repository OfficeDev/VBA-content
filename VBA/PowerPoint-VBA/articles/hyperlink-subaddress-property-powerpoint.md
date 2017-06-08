---
title: Hyperlink.SubAddress Property (PowerPoint)
keywords: vbapp10.chm526005
f1_keywords:
- vbapp10.chm526005
ms.prod: powerpoint
api_name:
- PowerPoint.Hyperlink.SubAddress
ms.assetid: f7b34b39-6e4c-5606-8b19-92ddc0dcede5
ms.date: 06/08/2017
---


# Hyperlink.SubAddress Property (PowerPoint)

Returns or sets the location within a document — such as a bookmark in a word document, a range in a Microsoft Office Excel worksheet, or a slide in a Microsoft PowerPoint presentation — associated with the specified hyperlink. Read/write.


## Syntax

 _expression_. **SubAddress**

 _expression_ A variable that represents a **Hyperlink** object.


### Return Value

String


## Example

This example sets shape one on slide one in the active presentation to jump to the slide named "Last Quarter" in Latest Figures.ppt when the shape is clicked during a slide show.


```vb
With ActivePresentation.Slides(1).Shapes(1) _
        .ActionSettings(ppMouseClick)
    .Action = ppActionHyperlink
    With .Hyperlink
        .Address = "c:\sales\latest figures.ppt"
        .SubAddress = "last quarter"
    End With
End With
```

This example sets shape one on slide one in the active presentation to jump to range A1:B10 in Latest.xls when the shape is clicked during a slide show.




```vb
With ActivePresentation.Slides(1).Shapes(1) _
        .ActionSettings(ppMouseClick)
    .Action = ppActionHyperlink
    With .Hyperlink
        .Address = "c:\sales\latest.xls"
        .SubAddress = "A1:B10"
    End With
End With
```


## See also


#### Concepts


[Hyperlink Object](hyperlink-object-powerpoint.md)

