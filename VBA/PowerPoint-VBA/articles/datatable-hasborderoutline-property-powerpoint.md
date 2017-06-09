---
title: DataTable.HasBorderOutline Property (PowerPoint)
keywords: vbapp10.chm698004
f1_keywords:
- vbapp10.chm698004
ms.prod: powerpoint
api_name:
- PowerPoint.DataTable.HasBorderOutline
ms.assetid: 16d6da74-b2a3-814c-e6d5-5686f8a36935
ms.date: 06/08/2017
---


# DataTable.HasBorderOutline Property (PowerPoint)

 **True** if the chart data table has outline borders. Read/write **Boolean**.


## Syntax

 _expression_. **HasBorderOutline**

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

