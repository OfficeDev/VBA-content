---
title: Axis.ReversePlotOrder Property (PowerPoint)
keywords: vbapp10.chm682025
f1_keywords:
- vbapp10.chm682025
ms.prod: powerpoint
api_name:
- PowerPoint.Axis.ReversePlotOrder
ms.assetid: 630d989b-1f9b-5258-d0be-479f362d2c66
ms.date: 06/08/2017
---


# Axis.ReversePlotOrder Property (PowerPoint)

 **True** if Microsoft Word plots data points from last to first. Read/write **Boolean**.


## Syntax

 _expression_. **ReversePlotOrder**

 _expression_ A variable that represents an **[Axis](axis-object-powerpoint.md)** object.


## Remarks

You cannot use this property on radar charts.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example plots data points from last to first on the value axis for the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.Axes(xlValue).ReversePlotOrder = True

    End If

End With
```


## See also


#### Concepts


[Axis Object](axis-object-powerpoint.md)

