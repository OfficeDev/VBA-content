---
title: QueryTable.MaintainConnection Property (Excel)
keywords: vbaxl10.chm518117
f1_keywords:
- vbaxl10.chm518117
ms.prod: excel
api_name:
- Excel.QueryTable.MaintainConnection
ms.assetid: e27fcb2d-115c-37c2-ba70-3f4a01dbb8b2
ms.date: 06/08/2017
---


# QueryTable.MaintainConnection Property (Excel)

 **True** if the connection to the specified data source is maintained after the refresh and until the workbook is closed. The default value is **True** . Read/write **Boolean** .


## Syntax

 _expression_ . **MaintainConnection**

 _expression_ A variable that represents a **QueryTable** object.


## Remarks

You can set the  **MaintainConnection** property only if the **[QueryType](querytable-querytype-property-excel.md)** property of the query table or PivotTable cache is set to **xlOLEDBQuery** .

If you anticipate frequent queries to a server, setting this property to  **True** might improve performance by reducing reconnection time. Setting the property to **False** causes an open connection to be closed.

If you import data using the user interface, data from a Web query or a text query is imported as a  **[QueryTable](querytable-object-excel.md)** object, while all other external data is imported as a **[ListObject](listobject-object-excel.md)** object.

If you import data using the object model, data from a Web query or a text query must be imported as a  **QueryTable** , while all other external data can be imported as either a **ListObject** or a **QueryTable** .

You can use the  **[QueryTable](listobject-querytable-property-excel.md)** property of the **ListObject** to access the **MaintainConnection** property.


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


[QueryTable Object](querytable-object-excel.md)

