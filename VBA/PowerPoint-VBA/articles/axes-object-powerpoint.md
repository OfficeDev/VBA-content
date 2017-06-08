---
title: Axes Object (PowerPoint)
keywords: vbapp10.chm681000
f1_keywords:
- vbapp10.chm681000
ms.prod: powerpoint
api_name:
- PowerPoint.Axes
ms.assetid: 71f1e1fc-7086-a84e-1e05-6fa50597b49b
ms.date: 06/08/2017
---


# Axes Object (PowerPoint)

Represents a collection of all the  **[Axis](axis-object-powerpoint.md)** objects in the specified chart.


## Remarks

Use the  **[Axes](chart-axes-method-powerpoint.md)** method to return the **Axes** collection.

Use  **Axes** ( _Type_, _AxisGroup_ ), where _Type_ is the axis type and _AxisGroup_ is the axis group, to return an **Axes** collection that contains a single **Axis** object. _Type_ can be one of the following **[XlAxisType](xlaxistype-enumeration-powerpoint.md)** constants: **xlCategory**, **xlSeries**, or **xlValue**. _AxisGroup_ can be one of the following **[XlAxisGroup](xlaxisgroup-enumeration-powerpoint.md)** constants: **xlPrimary** or **xlSecondary**.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example displays the number of axes for the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        MsgBox .Chart.Axes.Count

    End If

End With
```




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the category axis title text for the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.Axes(xlCategory)

            .HasTitle = True

            .AxisTitle.Caption = "1994"

        End With

    End If

End With
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

