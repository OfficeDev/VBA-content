---
title: QueryTables Object (Excel)
keywords: vbaxl10.chm520072
f1_keywords:
- vbaxl10.chm520072
ms.prod: excel
api_name:
- Excel.QueryTables
ms.assetid: 93511da3-598e-0aa3-fbc3-14bebff8838f
ms.date: 06/08/2017
---


# QueryTables Object (Excel)

A collection of  **[QueryTable](querytable-object-excel.md)** objects.


## Remarks

 Each **QueryTable** object represents a worksheet table built from data returned from an external data source.


## Example

Use the  **[QueryTables](worksheet-querytables-property-excel.md)** property to return the **[QueryTables](querytables-object-excel.md)** collection. The following example displays the number of query tables on the active worksheet.


```
MsgBox ActiveSheet.QueryTables.Count
```

Use the  **[Add](querytables-add-method-excel.md)** method to create a new query table and add it to the **QueryTables** collection. The following example creates a new query table.




```
Dim qt As QueryTable 
sqlstring = "select 96Sales.totals from 96Sales where profit < 5" 
connstring = _ 
 "ODBC;DSN=96SalesData;UID=Rep21;PWD=NUyHwYQI;Database=96Sales" 
With ActiveSheet.QueryTables.Add(Connection:=connstring, _ 
 Destination:=Range("B1"), Sql:=sqlstring) 
 .Refresh 
End With
```


## Methods



|**Name**|
|:-----|
|[Add](querytables-add-method-excel.md)|
|[Item](querytables-item-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](querytables-application-property-excel.md)|
|[Count](querytables-count-property-excel.md)|
|[Creator](querytables-creator-property-excel.md)|
|[Parent](querytables-parent-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
