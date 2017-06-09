---
title: PivotCaches.Create Method (Excel)
keywords: vbaxl10.chm229078
f1_keywords:
- vbaxl10.chm229078
ms.prod: excel
api_name:
- Excel.PivotCaches.Create
ms.assetid: d26e6786-064a-174c-5b9f-79e85b34f59b
ms.date: 06/08/2017
---


# PivotCaches.Create Method (Excel)

Creates a new PivotCache.


## Syntax

 _expression_ . **Create**( **_SourceType_** , **_SourceData_** , **_Version_** )

 _expression_ A variable that represents a **PivotCaches** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SourceType_|Required| **XlPivotTableSourceType**| _SourceType_ can be one of these **xlPivotTableSourceType** constants: **xlConsolidation** , **xlDatabase** , or **xlExternal** .|
| _SourceData_|Optional| **Variant**|The data for the new PivotTable cache.|
| _Version_|Optional| **Variant**|Version of the PivotTable. The version can be one of the [xlPivotTableVersionList](xlpivottableversionlist-enumeration-excel.md) constants.|

### Return Value

PivotCache


## Remarks

The following two  **xlPivotTableSourceType** constants are not supported when creating a PivotCache using this method: **xlPivotTable** and **xlScenario** . A run-time error is returned if one of these two constants is supplied.

The  _SourceData_ argument is required if _SourceType_ isn't **xlExternal** . It should be passed a Range (when _SourceType_ is either **xlConsolidation** or **xlDatabase** ) or an Excel Workbook Connection object (when _SourceType_ is **xlExternal** ). When passing a Range, it is recommended to either use a string to specify the workbook, worksheet, and cell range, or set up a named range and pass the name as a string. Passing a **Range** object may cause "type mismatch" errors unexpectedly.

When not supplied, the version of the PivotTable will be  **xlPivotTableVersion12** . The use of the **xlPivotTableVersionCurrent** constant is not allowed and returns a run-time error if it is supplied.


## Example

The following code sample defines a connection and then creates a connection to a  **PivotCache** .


```vb
Workbooks("Book1").Connections.Add2 _
        "Target Connection Name", "", Array("OLEDB;Provider=MSOLAP.5;Integrated Security=SSPI;Persist Security Info=True;Data Source=##TargetServer##;Initial Catalog=Adventure Works DW", ""), "Adventure Works", 1
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlExternal, SourceData:=ActiveWorkbook.Connections("Target Connection Name"), _ Version:=xlPivotTableVersion15).CreatePivotChart(ChartDestination:="Sheet1").Select

```


## See also


#### Concepts


[PivotCaches Object](pivotcaches-object-excel.md)

