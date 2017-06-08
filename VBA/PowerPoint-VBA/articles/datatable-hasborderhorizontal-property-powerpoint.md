---
title: DataTable.HasBorderHorizontal Property (PowerPoint)
keywords: vbapp10.chm698002
f1_keywords:
- vbapp10.chm698002
ms.prod: powerpoint
api_name:
- PowerPoint.DataTable.HasBorderHorizontal
ms.assetid: 6fb381e0-f990-656d-89e7-cb2f43a84ece
ms.date: 06/08/2017
---


# DataTable.HasBorderHorizontal Property (PowerPoint)

 **True** if the chart data table has horizontal cell borders. Read/write **Boolean**.


## Syntax

 _expression_. **HasBorderHorizontal**

 _expression_ A variable that represents a **[DataTable](datatable-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example causes the data table for the first chart in the active document to be displayed with an outline border and no cell borders.




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


[DataTable Object](datatable-object-powerpoint.md)

