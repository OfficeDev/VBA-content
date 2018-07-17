---
title: Chart.HeightPercent Property (PowerPoint)
keywords: vbapp10.chm684034
f1_keywords:
- vbapp10.chm684034
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.HeightPercent
ms.assetid: 71b6b6e3-ab2c-4ba3-cbbe-940fcbfe7efa
ms.date: 06/08/2017
---


# Chart.HeightPercent Property (PowerPoint)

Returns or sets the height of a 3-D chart as a percentage of the chart width (from 5 through 500 percent). Read/write  **Long**.


## Syntax

 _expression_. **HeightPercent**

 _expression_ A variable that represents a **[Chart](chart-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the height of the first chart in the active document to 80 percent of its width. You should run the example on a 3-D chart.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.HeightPercent = 80

    End If

End With
```


## See also


#### Concepts


[Chart Object](chart-object-powerpoint.md)

