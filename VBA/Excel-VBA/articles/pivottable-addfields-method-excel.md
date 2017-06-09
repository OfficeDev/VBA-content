---
title: PivotTable.AddFields Method (Excel)
keywords: vbaxl10.chm235073
f1_keywords:
- vbaxl10.chm235073
ms.prod: excel
api_name:
- Excel.PivotTable.AddFields
ms.assetid: b0ce878e-05a9-5c9a-4400-a26ba7c7162e
ms.date: 06/08/2017
---


# PivotTable.AddFields Method (Excel)

Adds row, column, and page fields to a PivotTable report or PivotChart report.


## Syntax

 _expression_ . **AddFields**( **_RowFields_** , **_ColumnFields_** , **_PageFields_** , **_AddToTable_** )

 _expression_ A variable that represents a **PivotTable** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _RowFields_|Optional| **Variant**|Specifies a field name (or an array of field names) to be added as rows, or to be added to the category axis.|
| _ColumnFields_|Optional| **Variant**|Specifies a field name (or an array of field names) to be added as columns, or to be added to the series axis.|
| _PageFields_|Optional| **Variant**|Specifies a field name (or an array of field names) to be added as pages, or to be added to the page area.|
| _AddToTable_|Optional| **Variant**|Applies only to PivotTable reports.  **True** to add the specified fields to the report (none of the existing fields are replaced). **False** to replace existing fields with the new fields. The default value is **False** .|

### Return Value

Variant


## Remarks

You must specify one of the field arguments.

Field names specify the unique name returned by the  **[SourceName](pivotfield-sourcename-property-excel.md)** property of the **[PivotField](pivotfield-object-excel.md)** object.

This method is not available for OLAP data sources.


## Example

This example replaces the existing column fields in the first PivotTable report on Sheet1 with the Status and Closed_By fields.


```vb
Worksheets("Sheet1").PivotTables(1).AddFields _ 
 ColumnFields:=Array("Status", "Closed_By")
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

