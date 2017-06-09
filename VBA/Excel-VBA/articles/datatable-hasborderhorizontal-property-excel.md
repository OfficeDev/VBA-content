---
title: DataTable.HasBorderHorizontal Property (Excel)
keywords: vbaxl10.chm626074
f1_keywords:
- vbaxl10.chm626074
ms.prod: excel
api_name:
- Excel.DataTable.HasBorderHorizontal
ms.assetid: 9d0f17f2-7786-afd5-164b-c7c5a4bb06d2
ms.date: 06/08/2017
---


# DataTable.HasBorderHorizontal Property (Excel)

 **True** if the chart data table has horizontal cell borders. Read/write **Boolean** .


## Syntax

 _expression_ . **HasBorderHorizontal**

 _expression_ A variable that represents a **DataTable** object.


## Example

This example causes the embedded chart data table to be displayed with an outline border and no cell borders.


```vb
With Worksheets(1).ChartObjects(1).Chart 
 .HasDataTable = True 
 With .DataTable 
 .HasBorderHorizontal = False 
 .HasBorderVertical = False 
 .HasBorderOutline = True 
 End With 
End With
```


## See also


#### Concepts


[DataTable Object](datatable-object-excel.md)

