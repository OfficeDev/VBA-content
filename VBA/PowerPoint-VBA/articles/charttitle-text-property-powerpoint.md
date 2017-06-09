---
title: ChartTitle.Text Property (PowerPoint)
keywords: vbapp10.chm694008
f1_keywords:
- vbapp10.chm694008
ms.prod: powerpoint
api_name:
- PowerPoint.ChartTitle.Text
ms.assetid: 01ae345d-d87e-31f4-de5d-85878289ad20
ms.date: 06/08/2017
---


# ChartTitle.Text Property (PowerPoint)

Returns or sets the text for the specified object. Read/write  **String**.


## Syntax

 _expression_. **Text**

 _expression_ A variable that represents a **[ChartTitle](charttitle-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the text for the chart title of the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.HasTitle = True

        .Chart.ChartTitle.Text = "First Quarter Sales"

    End If

End With
```


## See also


#### Concepts


[ChartTitle Object](charttitle-object-powerpoint.md)

