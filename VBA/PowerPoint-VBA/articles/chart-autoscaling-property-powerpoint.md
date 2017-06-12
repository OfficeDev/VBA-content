---
title: Chart.AutoScaling Property (PowerPoint)
keywords: vbapp10.chm684015
f1_keywords:
- vbapp10.chm684015
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.AutoScaling
ms.assetid: 330a185a-713a-409a-704e-3b163394aa92
ms.date: 06/08/2017
---


# Chart.AutoScaling Property (PowerPoint)

 **True** if Microsoft Word scales a 3-D chart so that it is closer in size to the equivalent 2-D chart. The **[RightAngleAxes](chart-rightangleaxes-property-powerpoint.md)** property must be **True**. Read/write **Boolean**.


## Syntax

 _expression_. **AutoScaling**

 _expression_ A variable that represents a **[Chart](chart-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example automatically scales the first chart in the active document. The example should be run on a 3-D chart.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.RightAngleAxes = True

        .Chart.AutoScaling = True

    End If

End With
```


## See also


#### Concepts


[Chart Object](chart-object-powerpoint.md)

