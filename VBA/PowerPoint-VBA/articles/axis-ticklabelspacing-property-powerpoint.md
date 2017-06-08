---
title: Axis.TickLabelSpacing Property (PowerPoint)
keywords: vbapp10.chm682030
f1_keywords:
- vbapp10.chm682030
ms.prod: powerpoint
api_name:
- PowerPoint.Axis.TickLabelSpacing
ms.assetid: 9a6694cb-bb6c-fc5d-a2a3-656327121581
ms.date: 06/08/2017
---


# Axis.TickLabelSpacing Property (PowerPoint)

Returns or sets the number of categories or series between tick-mark labels. Read/write  **Long**.


## Syntax

 _expression_. **TickLabelSpacing**

 _expression_ A variable that represents an **[Axis](axis-object-powerpoint.md)** object.


## Remarks

This property applies only to category and series axes. It can be a value from 1 through 31999. 

Tick-mark label spacing on the value axis is always calculated by Microsoft Word.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the number of categories between tick-mark labels on the category axis of the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.Axes(xlCategory).TickLabelSpacing = 10

    End If

End With
```


## See also


#### Concepts


[Axis Object](axis-object-powerpoint.md)

