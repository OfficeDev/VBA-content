---
title: Floor Object (PowerPoint)
keywords: vbapp10.chm703000
f1_keywords:
- vbapp10.chm703000
ms.prod: powerpoint
api_name:
- PowerPoint.Floor
ms.assetid: ed9ff3d1-8001-840c-d26e-7513ebe73ae9
ms.date: 06/08/2017
---


# Floor Object (PowerPoint)

Represents the floor of a 3-D chart.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

Use the  **[Floor](chart-floor-property-powerpoint.md)** property to return the **Floor** object. The following example sets the floor color for embedded chart one to cyan. The example will fail if the chart is not a 3-D chart.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.Floor.Interior.Color = RGB(0, 255, 255)

    End If

End With


```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

