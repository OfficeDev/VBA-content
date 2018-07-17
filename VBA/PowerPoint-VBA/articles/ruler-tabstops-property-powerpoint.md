---
title: Ruler.TabStops Property (PowerPoint)
keywords: vbapp10.chm570003
f1_keywords:
- vbapp10.chm570003
ms.prod: powerpoint
api_name:
- PowerPoint.Ruler.TabStops
ms.assetid: 11cc74dc-8efe-3327-87a1-0880e925040d
ms.date: 06/08/2017
---


# Ruler.TabStops Property (PowerPoint)

Returns a  **[TabStops](tabstops-object-powerpoint.md)** collection that represents the tab stops for the specified text. Read-only.


## Syntax

 _expression_. **TabStops**

 _expression_ A variable that represents a **Ruler** object.


### Return Value

TabStops


## Example

This example adds a slide with two text columns to the active presentation, sets a left-aligned tab stop for the title on the new slide, aligns the title box to the left, and assigns title text utilizing the tab stop just created.


```vb
With Application.ActivePresentation.Slides _
        .Add(2, ppLayoutTwoColumnText).Shapes

    With .Title.TextFrame
        With .Ruler
            .Levels(1).FirstMargin = 0
            .TabStops.Add ppTabStopLeft, 310
        End With
        .TextRange.ParagraphFormat.Alignment = ppAlignLeft
        .TextRange = "first column" + Chr(9) + "second column"
    End With

End With
```


## See also


#### Concepts


[Ruler Object](ruler-object-powerpoint.md)

