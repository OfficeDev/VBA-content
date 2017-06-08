---
title: TextFrame2.HorizontalAnchor Property (PowerPoint)
keywords: vbapp10.chm678007
f1_keywords:
- vbapp10.chm678007
ms.prod: powerpoint
api_name:
- PowerPoint.TextFrame2.HorizontalAnchor
ms.assetid: 17d27713-15c9-d846-f847-96e62768fafb
ms.date: 06/08/2017
---


# TextFrame2.HorizontalAnchor Property (PowerPoint)

 Returns or sets the horizontal alignment of text in a text frame. Read/write.


## Syntax

 _expression_. **HorizontalAnchor**

 _expression_ An expression that returns a **TextFrame2** object.


### Return Value

MsoHorizontalanchor


## Remarks

The value of the  **HorizontalAnchor** property can be one of these **MsoHorizontalAnchor** constants.


||
|:-----|
|**msoAnchorNone**|
|**msoHorizontalAnchorMixed**|
|**msoAnchorCenter**|

## Example

The following example shows how to set the alignment for shape one on slide one to top center.


```vb
Public Sub HorizontalAnchor_Example()



    With ActivePresentation.Slides(1).Shapes(1)

        .TextFrame2.HorizontalAnchor = msoAnchorCenter

        .TextFrame2.VerticalAnchor = msoAnchorTop

    End With

    

End Sub
```


## See also


#### Concepts


[TextFrame2 Object](textframe2-object-powerpoint.md)

