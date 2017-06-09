---
title: PivotCache Object (Excel)
keywords: vbaxl10.chm226072
f1_keywords:
- vbaxl10.chm226072
ms.prod: excel
api_name:
- Excel.PivotCache
ms.assetid: c3d84ef1-f9e6-b1bc-cbf0-3ba8dfe17439
ms.date: 06/08/2017
---


# PivotCache Object (Excel)

Represents the memory cache for a PivotTable report.


## Remarks

 The **PivotCache** object is a member of the **[PivotCaches](pivotcaches-object-excel.md)** collection.


## Example

Use the  **[PivotCache](pivottable-pivotcache-method-excel.md)** method to return a **PivotCache** object for a PivotTable report (each report has only one cache). The following example causes the first PivotTable report on the first worksheet to refresh itself whenever its file is opened.


```
Worksheets(1).PivotTables(1).PivotCache.RefreshOnFileOpen = True
```

Use  **[PivotCaches](workbook-pivotcaches-method-excel.md)** ( _index_ ), where _index_ is the PivotTable cache number, to return a single **PivotCache** object from the **PivotCaches** collection for a workbook. The following example refreshes cache one.




```
ActiveWorkbook.PivotCaches(1).Refresh
```


## Methods



|**Name**|
|:-----|
|[CreatePivotChart](pivotcache-createpivotchart-method-excel.md)|
|[CreatePivotTable](pivotcache-createpivottable-method-excel.md)|
|[MakeConnection](pivotcache-makeconnection-method-excel.md)|
|[Refresh](pivotcache-refresh-method-excel.md)|
|[ResetTimer](pivotcache-resettimer-method-excel.md)|
|[SaveAsODC](pivotcache-saveasodc-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[ADOConnection](pivotcache-adoconnection-property-excel.md)|
|[Application](pivotcache-application-property-excel.md)|
|[BackgroundQuery](pivotcache-backgroundquery-property-excel.md)|
|[CommandText](pivotcache-commandtext-property-excel.md)|
|[CommandType](pivotcache-commandtype-property-excel.md)|
|[Connection](pivotcache-connection-property-excel.md)|
|[Creator](pivotcache-creator-property-excel.md)|
|[EnableRefresh](pivotcache-enablerefresh-property-excel.md)|
|[Index](pivotcache-index-property-excel.md)|
|[IsConnected](pivotcache-isconnected-property-excel.md)|
|[LocalConnection](pivotcache-localconnection-property-excel.md)|
|[MaintainConnection](pivotcache-maintainconnection-property-excel.md)|
|[MemoryUsed](pivotcache-memoryused-property-excel.md)|
|[MissingItemsLimit](pivotcache-missingitemslimit-property-excel.md)|
|[OLAP](pivotcache-olap-property-excel.md)|
|[OptimizeCache](pivotcache-optimizecache-property-excel.md)|
|[Parent](pivotcache-parent-property-excel.md)|
|[QueryType](pivotcache-querytype-property-excel.md)|
|[RecordCount](pivotcache-recordcount-property-excel.md)|
|[Recordset](pivotcache-recordset-property-excel.md)|
|[RefreshDate](pivotcache-refreshdate-property-excel.md)|
|[RefreshName](pivotcache-refreshname-property-excel.md)|
|[RefreshOnFileOpen](pivotcache-refreshonfileopen-property-excel.md)|
|[RefreshPeriod](pivotcache-refreshperiod-property-excel.md)|
|[RobustConnect](pivotcache-robustconnect-property-excel.md)|
|[SavePassword](pivotcache-savepassword-property-excel.md)|
|[SourceConnectionFile](pivotcache-sourceconnectionfile-property-excel.md)|
|[SourceData](pivotcache-sourcedata-property-excel.md)|
|[SourceDataFile](pivotcache-sourcedatafile-property-excel.md)|
|[SourceType](pivotcache-sourcetype-property-excel.md)|
|[UpgradeOnRefresh](pivotcache-upgradeonrefresh-property-excel.md)|
|[UseLocalConnection](pivotcache-uselocalconnection-property-excel.md)|
|[Version](pivotcache-version-property-excel.md)|
|[WorkbookConnection](pivotcache-workbookconnection-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
