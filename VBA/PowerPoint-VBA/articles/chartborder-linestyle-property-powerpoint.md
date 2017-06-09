---
title: ChartBorder.LineStyle Property (PowerPoint)
keywords: vbapp10.chm685003
f1_keywords:
- vbapp10.chm685003
ms.prod: powerpoint
api_name:
- PowerPoint.ChartBorder.LineStyle
ms.assetid: 97ec4f20-72a4-b0a9-d875-c0ae0c492b1e
ms.date: 06/08/2017
---


# ChartBorder.LineStyle Property (PowerPoint)

Returns or sets the line style for the border. Read/write  **[XlLineStyle](xllinestyle-enumeration-powerpoint.md)**, **xlGray25**, **xlGray50**, **xlGray75**, or **xlAutomatic**.


## Syntax

 _expression_. **LineStyle**

 _expression_ A variable that represents a **[ChartBorder](chartborder-object-powerpoint.md)** object.


## Remarks

The  **xlDouble** and **xlSlantDashDot** constants of the **XlLineStyle** enumeration do not apply to charts.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example puts a border around the chart area and the plot area of the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart

            .ChartArea.Border.LineStyle = xlDashDot

            With .PlotArea.Border

                .LineStyle = xlDashDotDot

                .Weight = xlThick

            End With

        End With

    End If

End With


```


## See also


#### Concepts


[ChartBorder Object](chartborder-object-powerpoint.md)

