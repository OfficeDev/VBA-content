---
title: ChartData.IsLinked Property (PowerPoint)
keywords: vbapp10.chm689003
f1_keywords:
- vbapp10.chm689003
ms.prod: powerpoint
api_name:
- PowerPoint.ChartData.IsLinked
ms.assetid: 038ed026-a14c-2c5c-3f2e-c931fa9840b0
ms.date: 06/08/2017
---


# ChartData.IsLinked Property (PowerPoint)

 **True** if the data for the chart is linked to an external Microsoft Excel workbook. Read-only **Boolean**.


## Syntax

 _expression_. **IsLinked**

 _expression_ A variable that represents a **[ChartData](chartdata-object-powerpoint.md)** object.


## Remarks

Using the  **[BreakLink](chartdata-breaklink-method-powerpoint.md)** method to remove the link to an Excel workbook sets this property to **False**.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example verifies whether the data for the first chart in the active document is linked to an external Excel workbook. If the data for the chart is linked, the example then uses the  **BreakLink** method to remove the link. If the data for the chart is not linked, the example uses the **[Activate](chartdata-activate-method-powerpoint.md)** method to display the embedded data for the chart.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.ChartData

            If .IsLinked Then

                .BreakLink

            Else

                .Activate

            End If

        End With

    End If

End With
```


## See also


#### Concepts


[ChartData Object](chartdata-object-powerpoint.md)

