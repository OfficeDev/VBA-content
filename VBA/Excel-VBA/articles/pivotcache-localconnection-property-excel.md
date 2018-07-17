---
title: PivotCache.LocalConnection Property (Excel)
keywords: vbaxl10.chm227094
f1_keywords:
- vbaxl10.chm227094
ms.prod: excel
api_name:
- Excel.PivotCache.LocalConnection
ms.assetid: 3afee878-3c05-6b05-4770-e10e4c6f9375
ms.date: 06/08/2017
---


# PivotCache.LocalConnection Property (Excel)

Returns or sets the connection string to an offline cube file. Read/write  **String** .


## Syntax

 _expression_ . **LocalConnection**

 _expression_ A variable that represents a **PivotCache** object.


## Remarks

For a non-OLAP data source, the value of the  **LocalConnection** property is an empty string, and the **[UseLocalConnection](pivotcache-uselocalconnection-property-excel.md)** property is set to **False** .

Setting the  **LocalConnection** property does not immediately initiate the connection to the data source. You must first use the **Refresh** method to make the connection and retrieve the data.

The value of the  **LocalConnection** property is used if the **UseLocalConnection** property is set to **True** . If the **UseLocalConnection** property is set to **False** , the **Connection** property specifies the connection string for query tables based on sources other than local cube files.

For more information about the syntax for connection strings, see the Help topic for the  **Add** method of the PivotTable collection.


## Example

This example sets the connection string of the first PivotTable cache to reference an offline cube file.


```vb
With ActiveWorkbook.PivotCaches(1) 
 .LocalConnection = _ 
 "OLEDB;Provider=MSOLAP;Data Source=C:\Data\DataCube.cub" 
 .UseLocalConnection = True 
End With 

```


## See also


#### Concepts


[PivotCache Object](pivotcache-object-excel.md)

