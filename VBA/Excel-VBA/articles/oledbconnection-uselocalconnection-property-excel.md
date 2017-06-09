---
title: OLEDBConnection.UseLocalConnection Property (Excel)
keywords: vbaxl10.chm794094
f1_keywords:
- vbaxl10.chm794094
ms.prod: excel
api_name:
- Excel.OLEDBConnection.UseLocalConnection
ms.assetid: b346933c-17cd-ef11-6070-ee840c8d7c0a
ms.date: 06/08/2017
---


# OLEDBConnection.UseLocalConnection Property (Excel)

 **True** if the **[LocalConnection](oledbconnection-localconnection-property-excel.md)** property is used to specify the string that enables Microsoft Excel to connect to a data source. **False** if the connection string specified by the **[Connection](oledbconnection-connection-property-excel.md)** property is used. Read/write **Boolean** .


## Syntax

 _expression_ . **UseLocalConnection**

 _expression_ A variable that represents an **OLEDBConnection** object.


## Remarks

This example sets the connection string of the first PivotTable cache to reference an offline cube file.


## Example


```vb
With ActiveWorkbook.PivotCaches(1) 
 .LocalConnection = _ 
 "OLEDB;Provider=MSOLAP;Data Source=C:\Data\DataCube.cub" 
 .UseLocalConnection = True 
End With 

```


## See also


#### Concepts


[OLEDBConnection Object](oledbconnection-object-excel.md)

