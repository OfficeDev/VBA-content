---
title: AxisTitle Object (PowerPoint)
keywords: vbapp10.chm683000
f1_keywords:
- vbapp10.chm683000
ms.prod: powerpoint
api_name:
- PowerPoint.AxisTitle
ms.assetid: 8eddc95c-2353-43fa-c055-ee76de28009d
ms.date: 06/08/2017
---


# AxisTitle Object (PowerPoint)

Represents a chart axis title.


## Remarks

Use the  **[AxisTitle](axis-axistitle-property-powerpoint.md)** property to return an **AxisTitle** object.

The  **AxisTitle** object does not exist and cannot be used unless the **[HasTitle](axis-hastitle-property-powerpoint.md)** property for the axis is **True**.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the caption, sets the font to Bookman 10 point, and formats the word "millions" as italic for the axis title of the value axis for the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.Axes(xlValue)

            .HasTitle = True

            With .AxisTitle

                .Caption = "Revenue (millions)"

                .Font.Name = "bookman"

                .Font.Size = 10

                .Characters(10, 8).Font.Italic = True

            End With

        End With

    End If

End With


```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

