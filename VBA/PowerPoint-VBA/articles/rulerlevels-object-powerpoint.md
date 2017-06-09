---
title: RulerLevels Object (PowerPoint)
keywords: vbapp10.chm571000
f1_keywords:
- vbapp10.chm571000
ms.prod: powerpoint
api_name:
- PowerPoint.RulerLevels
ms.assetid: 890f4bee-c48a-be48-2cac-b73736a5bdf0
ms.date: 06/08/2017
---


# RulerLevels Object (PowerPoint)

A collection of all the  **[RulerLevel](rulerlevel-object-powerpoint.md)** objects on the specified ruler.


## Remarks

Each  **RulerLevel** object represents the first-line and left indent for text at a particular outline level. This collection always contains five members — one for each of the available outline levels.


## Example

Use the [Levels](ruler-levels-property-powerpoint.md)property to return the  **RulerLevels** collection. The following example sets the margins for the five outline levels in body text in the active presentation.


```vb
With ActivePresentation.SlideMaster.TextStyles(ppBodyStyle).Ruler

    .Levels(1).FirstMargin = 0

    .Levels(1).LeftMargin = 40

    .Levels(2).FirstMargin = 60

    .Levels(2).LeftMargin = 100

    .Levels(3).FirstMargin = 120

    .Levels(3).LeftMargin = 160

    .Levels(4).FirstMargin = 180

    .Levels(4).LeftMargin = 220

    .Levels(5).FirstMargin = 240

    .Levels(5).LeftMargin = 280

End With
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

