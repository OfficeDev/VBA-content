---
title: ParagraphFormat.SpaceBefore Property (PowerPoint)
keywords: vbapp10.chm576008
f1_keywords:
- vbapp10.chm576008
ms.prod: powerpoint
api_name:
- PowerPoint.ParagraphFormat.SpaceBefore
ms.assetid: be73b3fe-4490-df58-57fd-47c51767b985
ms.date: 06/08/2017
---


# ParagraphFormat.SpaceBefore Property (PowerPoint)

Returns or sets the amount of space before the first line in each paragraph of the specified text, in points or lines. Read/write.


## Syntax

 _expression_. **SpaceBefore**

 _expression_ A variable that represents a **ParagraphFormat** object.


### Return Value

Single


## Example

This example sets the spacing before paragraphs to 6 points for the text in shape two on slide in the active presentation.


```vb
With Application.ActivePresentation.Slides(1).Shapes(2)

    With .TextFrame.TextRange.ParagraphFormat

        .LineRuleBefore = False

        .SpaceBefore = 6

    End With

End With
```


## See also


#### Concepts


[ParagraphFormat Object](paragraphformat-object-powerpoint.md)

