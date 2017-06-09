---
title: BulletFormat.RelativeSize Property (PowerPoint)
keywords: vbapp10.chm577005
f1_keywords:
- vbapp10.chm577005
ms.prod: powerpoint
api_name:
- PowerPoint.BulletFormat.RelativeSize
ms.assetid: ce90fbcb-9aa5-a286-1f91-f06a83351b97
ms.date: 06/08/2017
---


# BulletFormat.RelativeSize Property (PowerPoint)

Returns or sets the bullet size relative to the size of the first text character in the paragraph. Read/write.


## Syntax

 _expression_. **RelativeSize**

 _expression_ A variable that represents a **BulletFormat** object.


### Return Value

Single


## Remarks

The  **RelativeSize** property value can be a floating-point value from 0.25 through 4, indicating that the bullet size can be from 25 percent through 400 percent of the text-character size.


## Example

This example sets the formatting for the bullet in shape two on slide one in the active presentation. The size of the bullet is 125 percent of the size of the first text character in the paragraph.


```vb
With ActivePresentation.Slides(1).Shapes(2)

    With .TextFrame.TextRange.ParagraphFormat.Bullet

        .Visible = True

        .RelativeSize = 1.25

        .Character = 169

        With .Font

            .Name = "Symbol"

            .Color.RGB = RGB(255, 0, 0)

        End With

    End With

End With
```


## See also


#### Concepts


[BulletFormat Object](bulletformat-object-powerpoint.md)

