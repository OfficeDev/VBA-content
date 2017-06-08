---
title: ParagraphFormat.SpaceWithin Property (PowerPoint)
keywords: vbapp10.chm576010
f1_keywords:
- vbapp10.chm576010
ms.prod: powerpoint
api_name:
- PowerPoint.ParagraphFormat.SpaceWithin
ms.assetid: 523fa767-e5af-0d7f-d16a-b11dd7d3799d
ms.date: 06/08/2017
---


# ParagraphFormat.SpaceWithin Property (PowerPoint)

Returns or sets the amount of space between base lines in the specified text, in points or lines. Read/write.


## Syntax

 _expression_. **SpaceWithin**

 _expression_ A variable that represents a **ParagraphFormat** object.


### Return Value

Single


## Example

This example sets line spacing to 21 points for the text in shape two on slide two in the active presentation.


```vb
With Application.ActivePresentation.Slides(2).Shapes(2)

    With .TextFrame.TextRange.ParagraphFormat

        .LineRuleWithin = False

        .SpaceWithin = 21

    End With

End With
```


## See also


#### Concepts


[ParagraphFormat Object](paragraphformat-object-powerpoint.md)

