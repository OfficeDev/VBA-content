---
title: TextStyleLevel.ParagraphFormat Property (PowerPoint)
keywords: vbapp10.chm581003
f1_keywords:
- vbapp10.chm581003
ms.prod: powerpoint
api_name:
- PowerPoint.TextStyleLevel.ParagraphFormat
ms.assetid: bc49660b-7834-0c6c-230f-0d9d31543c71
ms.date: 06/08/2017
---


# TextStyleLevel.ParagraphFormat Property (PowerPoint)

Returns a  **[ParagraphFormat](paragraphformat-object-powerpoint.md)** object that represents paragraph formatting for the specified text. Read-only.


## Syntax

 _expression_. **ParagraphFormat**

 _expression_ A variable that represents a **TextStyleLevel** object.


### Return Value

ParagraphFormat


## Example

This example sets the line spacing before, within, and after each paragraph in shape two on slide one in the active presentation.


```vb
With Application.ActivePresentation.Slides(2).Shapes(2)

    With .TextFrame.TextRange.ParagraphFormat

        .LineRuleWithin = msoTrue

        .SpaceWithin = 1.4

        .LineRuleBefore = msoTrue

        .SpaceBefore = 0.25

        .LineRuleAfter = msoTrue

        .SpaceAfter = 0.75

    End With

End With
```


## See also


#### Concepts


[TextStyleLevel Object](textstylelevel-object-powerpoint.md)

