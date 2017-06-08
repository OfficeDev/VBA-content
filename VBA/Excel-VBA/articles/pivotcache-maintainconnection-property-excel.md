---
title: PivotCache.MaintainConnection Property (Excel)
keywords: vbaxl10.chm227090
f1_keywords:
- vbaxl10.chm227090
ms.prod: excel
api_name:
- Excel.PivotCache.MaintainConnection
ms.assetid: 1fba45e7-0059-26d1-1433-631ee08c0dd0
ms.date: 06/08/2017
---


# PivotCache.MaintainConnection Property (Excel)

 **True** if the connection to the specified data source is maintained after the refresh and until the workbook is closed. The default value is **True** . Read/write **Boolean** .


## Syntax

 _expression_ . **MaintainConnection**

 _expression_ A variable that represents a **PivotCache** object.


## Remarks

You can set the  **MaintainConnection** property only if the **[QueryType](querytable-querytype-property-excel.md)** property of the query table or PivotTable cache is set to **xlOLEDBQuery** .

If you anticipate frequent queries to a server, setting this property to  **True** might improve performance by reducing reconnection time. Setting the property to **False** causes an open connection to be closed.


## Example

This example creates a new PivotTable cache based on an OLAP provider, and then it creates a new PivotTable report based on the cache, at cell A3 on the active worksheet. The example terminates the connection after the initial refresh.


```vb
With ActiveWorkbook.PivotCaches.Add(SourceType:=xlExternal) 
 .Connection = _ 
 "OLEDB;Provider=MSOLAP;Location=srvdata;Initial Catalog=National" 
 .MaintainConnection = False 
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


[PivotCache Object](pivotcache-object-excel.md)

