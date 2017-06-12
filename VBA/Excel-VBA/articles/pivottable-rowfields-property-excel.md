---
title: PivotTable.RowFields Property (Excel)
keywords: vbaxl10.chm235093
f1_keywords:
- vbaxl10.chm235093
ms.prod: excel
api_name:
- Excel.PivotTable.RowFields
ms.assetid: 3976d5ec-b248-55f5-659d-2671af3f3bfd
ms.date: 06/08/2017
---


# PivotTable.RowFields Property (Excel)

Returns an object that represents either a single field in a PivotTable report (a  **[PivotField](pivotfield-object-excel.md)** object) or a collection of all the fields (a **[PivotFields](pivotfields-object-excel.md)** object) that are currently showing as row fields. Read-only.


## Syntax

 _expression_ . **RowFields**( **_Index_** )

 _expression_ A variable that represents a **PivotTable** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|The name or number of the field to be returned (can be an array to specify more than one field).|

## Example

This example adds the PivotTable report?s row field names to a list on a new worksheet.


```vb
Set nwSheet = Worksheets.Add 
nwSheet.Activate 
Set pvtTable = Worksheets("Sheet2").Range("A1").PivotTable 
rw = 0 
For Each pvtField In pvtTable.RowFields 
 rw = rw + 1 
 nwSheet.Cells(rw, 1).Value = pvtField.Name 
Next pvtField
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

