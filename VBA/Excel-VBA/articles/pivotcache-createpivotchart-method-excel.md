---
title: PivotCache.CreatePivotChart Method (Excel)
keywords: vbaxl10.chm227110
f1_keywords:
- vbaxl10.chm227110
ms.prod: excel
ms.assetid: 5aeb9a16-2cf8-3525-12b0-0b6e3d3ddf1a
ms.date: 06/08/2017
---


# PivotCache.CreatePivotChart Method (Excel)

Creates a standalone PivotChart from a [PivotCache Object (Excel)](pivotcache-object-excel.md) object. A[Shape Object (Excel)](shape-object-excel.md) object is returned.


## Syntax

 _expression_ . **CreatePivotChart**_(ChartDestination,_ _XlChartType,_ _Left,_ _Top,_ _Width,_ _Height)_

 _expression_ A variable that represents a[PivotCache Object (Excel)](pivotcache-object-excel.md) object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ChartDestination_|Required|VARIANT|The Destination worksheet|
| _XlChartType_|Optional|VARIANT|The type of chart|
| _Left_|Optional|VARIANT|The distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).|
| _Top_|Optional|VARIANT|The distance, in points, from the top edge of the topmost shape in the shape range to the top edge of the worksheet.|
| _Width_|Optional|VARIANT|The width, in points, of the object.|
| _Height_|Optional|VARIANT|The height, in points, of the object.|

### Return value

[Shape Object (Excel)](shape-object-excel.md) object


## Remarks

If the  **PivotCache** object that the method is called from has no attached PivotTable:


- A workbook-level PivotTable is created from the existing PivotCache.
    
- A standalone PivotChart will be created with a reference to the newly created PivotTable.
    
If the PivotCache already has an associated PivotTable:


- The PivotCache is cloned
    
- A new workbook-level PivotTable is created based on the cloned PivotCache.
    
- A standalone PivotChart is created with a reference to the new workbook-level PivotTable.
    

## Example

The following code creates a decoupled PivotChart from a PivotCache object.


```vb
Workbooks("Book1").Connections.Add _
     "cubes4 Adventure Works DW 2008 Special Char Adventure Works", "", Array( _
     "OLEDB;Provider=MSOLAP.4;Integrated Security=SSPI;Persist Security Info=True;Data Source=<server name here >;Initial Catalog=Adventure Works DW 2008" _
     , " Special Char"), Array("Adventure Works"), 1
   ActiveWorkbook.PivotCaches.Create(SourceType:=xlExternal, SourceData:= _
     ActiveWorkbook.Connections( _
     "cubes4 Adventure Works DW 2008 Special Char Adventure Works"), Version:= _
     xlPivotTableVersion14).CreatePivotChart(ChartDestination:="Sheet1").Select

   ActiveChart.ChartType = xlColumnClustered
```


## See also


#### Concepts


[PivotCache Object](pivotcache-object-excel.md)

