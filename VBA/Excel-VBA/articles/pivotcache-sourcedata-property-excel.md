---
title: PivotCache.SourceData Property (Excel)
keywords: vbaxl10.chm227086
f1_keywords:
- vbaxl10.chm227086
ms.prod: excel
api_name:
- Excel.PivotCache.SourceData
ms.assetid: 5a172543-3a06-9db0-7edc-0cf2aa7af114
ms.date: 06/08/2017
---


# PivotCache.SourceData Property (Excel)

Returns the data source for the PivotTable report, as shown in the following table. Read-write  **Variant** .


## Syntax

 _expression_ . **SourceData**

 _expression_ A variable that represents a **PivotCache** object.


## Remarks



|**Data source**|**Return value**|
|:-----|:-----|
|Microsoft Excel list or database|The cell reference, as text.|
|External data source|An array. Each row consists of an SQL connection string with the remaining elements as the query string, broken down into 255-character segments.|
|Multiple consolidation ranges|A two-dimensional array. Each row consists of a reference and its associated page field items.|
|Another PivotTable report|One of the above three kinds of information.|
This property is not available for OLE DB data sources.


## Example

Assume that you used an external data source to create a PivotTable report on Sheet1. This example inserts the SQL connection string and query string into a new worksheet.


```vb
Set newSheet = ActiveWorkbook.Worksheets.Add 
sdArray = Worksheets("Sheet1").UsedRange.PivotTable.SourceData 
For i = LBound(sdArray) To UBound(sdArray) 
 newSheet.Cells(i, 1) = sdArray(i) 
Next i 

```


## See also


#### Concepts


[PivotCache Object](pivotcache-object-excel.md)

