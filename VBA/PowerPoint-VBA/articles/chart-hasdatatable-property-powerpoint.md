---
title: Chart.HasDataTable Property (PowerPoint)
keywords: vbapp10.chm66932
f1_keywords:
- vbapp10.chm66932
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.HasDataTable
ms.assetid: 6864181a-da77-9da5-adad-008ecc5c8f7f
ms.date: 06/08/2017
---


# Chart.HasDataTable Property (PowerPoint)

 **True** if the chart has a data table. Read/write **Boolean**.


## Syntax

 _expression_. **HasDataTable**

 _expression_ A variable that represents a **[Chart](chart-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example causes the embedded chart data table to be displayed with an outline border and no cell borders.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart

            .HasDataTable = True

            With .DataTable

                .HasBorderHorizontal = False

                .HasBorderVertical = False

                .HasBorderOutline = True

            End With

        End With

    End If

End With


```


## See also


#### Concepts


[Chart Object](chart-object-powerpoint.md)

