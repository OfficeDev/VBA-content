---
title: PivotCache.UseLocalConnection Property (Excel)
keywords: vbaxl10.chm227096
f1_keywords:
- vbaxl10.chm227096
ms.prod: excel
api_name:
- Excel.PivotCache.UseLocalConnection
ms.assetid: ce54adf2-22f3-f4dc-8b97-276d6ca53478
ms.date: 06/08/2017
---


# PivotCache.UseLocalConnection Property (Excel)

Returns  **True** if the **[LocalConnection](pivotcache-localconnection-property-excel.md)** property is used to specify the string that enables Microsoft Excel to connect to a data source. Returns **False** if the connection string specified by the **Connection** property is used. Read/write **Boolean** .


## Syntax

 _expression_ . **UseLocalConnection**

 _expression_ A variable that represents a **PivotCache** object.


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

