---
title: ChartBorder.Weight Property (PowerPoint)
keywords: vbapp10.chm685004
f1_keywords:
- vbapp10.chm685004
ms.prod: powerpoint
api_name:
- PowerPoint.ChartBorder.Weight
ms.assetid: 71750026-1df0-1a1b-bb43-b0c6891d66be
ms.date: 06/08/2017
---


# ChartBorder.Weight Property (PowerPoint)

Returns or sets the weight of the border. Read/write  **[XlBorderWeight](xlborderweight-enumeration-powerpoint.md)**.


## Syntax

 _expression_. **Weight**

 _expression_ A variable that represents a **[ChartBorder](chartborder-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the border weight for the value axis of the first chart in the active document to medium.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.Axes(xlValue).Border.Weight = xlMedium

    End If

End With
```


## See also


#### Concepts


[ChartBorder Object](chartborder-object-powerpoint.md)

