---
title: OLEDBConnection.Connection Property (Excel)
keywords: vbaxl10.chm794078
f1_keywords:
- vbaxl10.chm794078
ms.prod: excel
api_name:
- Excel.OLEDBConnection.Connection
ms.assetid: 03b83f0e-1a16-f44e-0a89-27742b733e05
ms.date: 06/08/2017
---


# OLEDBConnection.Connection Property (Excel)

Returns or sets a string that contains OLE DB settings that enable Microsoft Excel to connect to an OLE DB data source. Read/write  **Variant** .


## Syntax

 _expression_ . **Connection**

 _expression_ A variable that represents an **OLEDBConnection** object.


## Remarks

Setting the  **Connection** property does not immediately initiate the connection to the data source. You must use the **[Refresh](oledbconnection-refresh-method-excel.md)** method to make the connection and retrieve the data. When using an offline cube file, set the **[UseLocalConnection](oledbconnection-uselocalconnection-property-excel.md)** property to **True** and use the **[LocalConnection](oledbconnection-localconnection-property-excel.md)** property instead of the **Connection** property.


## Example

This example creates a PivotTable cache based on an OLAP provider, and then it creates a PivotTable report based on the cache, at cell A3 on the active worksheet.


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


[OLEDBConnection Object](oledbconnection-object-excel.md)

