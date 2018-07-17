---
title: PivotCache.Connection Property (Excel)
keywords: vbaxl10.chm227074
f1_keywords:
- vbaxl10.chm227074
ms.prod: excel
api_name:
- Excel.PivotCache.Connection
ms.assetid: 5d4b07f2-dad9-4c90-ec92-094dac95a086
ms.date: 06/08/2017
---


# PivotCache.Connection Property (Excel)

Returns or sets a string that contains one of the following: OLE DB settings that enable Microsoft Excel to connect to an OLE DB data source; ODBC settings that enable Microsoft Excel to connect to an ODBC data source; a URL that enables Microsoft Excel to connect to a Web data source; the path to and file name of a text file, or the path to and file name of a file that specifies a database or Web query. Read/write  **Variant** .


## Syntax

 _expression_ . **Connection**

 _expression_ An expression that returns a **PivotCache** object.


## Remarks

When using an offline cube file, set the  **[UseLocalConnection](pivotcache-uselocalconnection-property-excel.md)** property to **True** and use the **[LocalConnection](pivotcache-localconnection-property-excel.md)** property instead of the **Connection** property.

Alternatively, you may choose to access a data source directly by using the Microsoft ActiveX Data Objects (ADO) library instead.


## Example

This example creates a new PivotTable cache based on an OLAP provider, and then it creates a new PivotTable report based on the cache, at cell A3 on the active worksheet.


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


[PivotCache Object](pivotcache-object-excel.md)

