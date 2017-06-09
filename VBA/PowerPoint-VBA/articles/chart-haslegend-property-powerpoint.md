---
title: Chart.HasLegend Property (PowerPoint)
keywords: vbapp10.chm684032
f1_keywords:
- vbapp10.chm684032
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.HasLegend
ms.assetid: 084f7de3-b0ed-d7b3-3b24-465e74afa167
ms.date: 06/08/2017
---


# Chart.HasLegend Property (PowerPoint)

 **True** if the chart has a legend. Read/write **Boolean**.


## Syntax

 _expression_. **HasLegend**

 _expression_ A variable that represents a **[Chart](chart-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example enables the legend for the first chart in the active document and then sets the legend font color to blue.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart

            .HasLegend = True

            .Legend.Font.ColorIndex = 5

        End With

    End If

End With
```


## See also


#### Concepts


[Chart Object](chart-object-powerpoint.md)

