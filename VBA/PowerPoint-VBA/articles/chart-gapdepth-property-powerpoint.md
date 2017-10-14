---
title: Chart.GapDepth Property (PowerPoint)
keywords: vbapp10.chm684030
f1_keywords:
- vbapp10.chm684030
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.GapDepth
ms.assetid: e25fdb39-735f-6158-2628-b0696c08b733
ms.date: 06/08/2017
---


# Chart.GapDepth Property (PowerPoint)

Returns or sets the distance, as a percentage of the marker width, between the data series in a 3-D chart. Read/write  **Long**.


## Syntax

 _expression_. **GapDepth**

 _expression_ A variable that represents a **[Chart](chart-object-powerpoint.md)** object.


## Remarks

The value of this property must be between 0 and 500. 


 **Note**  This property applies only to 3-D charts.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the distance between the data series for the first chart in the active document to 200 percent of the marker width. You should run the example on a 3-D chart (the  **GapDepth** property fails on 2-D charts).




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.GapDepth = 200

    End If

End With
```


## See also


#### Concepts


[Chart Object](chart-object-powerpoint.md)

