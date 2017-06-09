---
title: ChartGroup.BubbleScale Property (PowerPoint)
keywords: vbapp10.chm692008
f1_keywords:
- vbapp10.chm692008
ms.prod: powerpoint
api_name:
- PowerPoint.ChartGroup.BubbleScale
ms.assetid: ecc3f3e1-512c-cbd1-094a-337d5f2ba83f
ms.date: 06/08/2017
---


# ChartGroup.BubbleScale Property (PowerPoint)

Returns or sets the scale factor for bubbles in the specified chart group. Read/write  **Long**.


## Syntax

 _expression_. **BubbleScale**

 _expression_ A variable that represents a **[ChartGroup](chartgroup-object-powerpoint.md)** object.


## Remarks

You can set this property to an integer from 0 (zero) through 300, corresponding to a percentage of the default size. 


 **Note**  This property applies only to bubble charts.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the bubble size in the first chart group of the first chart in the active document to 200 percent of the default size.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.ChartGroups(1).BubbleScale = 200

    End If

End With
```


## See also


#### Concepts


[ChartGroup Object](chartgroup-object-powerpoint.md)

