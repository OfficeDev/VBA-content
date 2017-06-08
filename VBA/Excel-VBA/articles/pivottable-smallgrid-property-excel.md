---
title: PivotTable.SmallGrid Property (Excel)
keywords: vbaxl10.chm235134
f1_keywords:
- vbaxl10.chm235134
ms.prod: excel
api_name:
- Excel.PivotTable.SmallGrid
ms.assetid: ade36fce-e511-f95c-db92-e64271646687
ms.date: 06/08/2017
---


# PivotTable.SmallGrid Property (Excel)

 **True** if Microsoft Excel uses a grid that's two cells wide and two cells deep for a newly created PivotTable report. **False** if Excel uses a blank stencil outline. Read/write **Boolean** .


## Syntax

 _expression_ . **SmallGrid**

 _expression_ A variable that represents a **PivotTable** object.


## Remarks

You should use the stencil outline. The grid is provided only because it enables compatibility with earlier versions of Excel.


## Example

This example creates a new PivotTable cache based on an OLAP provider, and then it creates a new PivotTable report based on this cache, at cell A3 on the active worksheet. The example uses the stencil outline instead of the cell grid.


```vb
With ActiveWorkbook.PivotCaches.Add(SourceType:=xlExternal) 
 .Connection = _ 
 "OLEDB;Provider=MSOLAP;Location=srvdata;Initial Catalog=National" 
 .MaintainConnection = True 
 .CreatePivotTable TableDestination:=Range("A3"), _ 
 TableName:= "PivotTable1" 
End With 
With ActiveSheet.PivotTables("PivotTable1") 
 .SmallGrid = False 
 .PivotCache.RefreshPeriod = 0 
 With .CubeFields("[state]") 
 .Orientation = xlColumnField 
 .Position = 0 
 End With 
 With .CubeFields("[Measures].[Count Of au_id]") 
 .Orientation = xlDataField 
 .Position = 0 
 End With 
End With
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

