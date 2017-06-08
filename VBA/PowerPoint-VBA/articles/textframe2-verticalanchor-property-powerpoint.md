---
title: TextFrame2.VerticalAnchor Property (PowerPoint)
keywords: vbapp10.chm678008
f1_keywords:
- vbapp10.chm678008
ms.prod: powerpoint
api_name:
- PowerPoint.TextFrame2.VerticalAnchor
ms.assetid: e00b1b4b-c291-fb10-be85-49e84ab0b739
ms.date: 06/08/2017
---


# TextFrame2.VerticalAnchor Property (PowerPoint)

 Returns or sets the vertical alignment of text in a text frame. Read/write.


## Syntax

 _expression_. **VerticalAnchor**

 _expression_ An expression that returns a **TextFrame2** object.


## Remarks

The value of the  **VerticalAnchor** property can be one of these **MsoVerticalAnchor** constants.


||
|:-----|
|**msoAnchorBottom**|
|**msoAnchorMiddle**|
|**msoAnchorTop**|
|**msoVerticalAnchorMixed**|

## Example

The following example shows how to set the alignment for shape one on slide one to top center.


```vb
Public Sub VerticalAnchor_Example()



    With ActivePresentation.Slides(1).Shapes(1)

        .TextFrame2.HorizontalAnchor = msoAnchorCenter

        .TextFrame2.VerticalAnchor = msoAnchorTop

    End With

    

End Sub
```


## See also


#### Concepts


[TextFrame2 Object](textframe2-object-powerpoint.md)

