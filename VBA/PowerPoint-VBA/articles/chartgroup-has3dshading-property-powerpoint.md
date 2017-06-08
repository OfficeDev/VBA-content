---
title: ChartGroup.Has3DShading Property (PowerPoint)
keywords: vbapp10.chm692005
f1_keywords:
- vbapp10.chm692005
ms.prod: powerpoint
api_name:
- PowerPoint.ChartGroup.Has3DShading
ms.assetid: 6276bf7a-9d21-9eda-6ad9-6153c9a74948
ms.date: 06/08/2017
---


# ChartGroup.Has3DShading Property (PowerPoint)

 **True** if a chart group has three-dimensional shading. Read/write **Boolean**.


## Syntax

 _expression_. **Has3DShading**

 _expression_ A variable that represents a **[ChartGroup](chartgroup-object-powerpoint.md)** object.


## Remarks

Setting  **Has3DShading** to **False** removes the 3-D shading effect from the chart (rendering it as flat). Setting **Has3DShading** to **True** sets the chart content to the default.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example adds 3-D shading to the first chart group of the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.ChartGroups(1).Has3DShading = True

    End If

End With


```


## See also


#### Concepts


[ChartGroup Object](chartgroup-object-powerpoint.md)

