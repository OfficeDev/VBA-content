---
title: PivotTable.ColumnFields Property (Excel)
keywords: vbaxl10.chm235074
f1_keywords:
- vbaxl10.chm235074
ms.prod: excel
api_name:
- Excel.PivotTable.ColumnFields
ms.assetid: caae2016-e213-31f0-5ce7-fd8593ad4266
ms.date: 06/08/2017
---


# PivotTable.ColumnFields Property (Excel)

Returns an object that represents either a single PivotTable field (a  **[PivotField](pivotfield-object-excel.md)** object) or a collection of all the fields (a **[PivotFields](pivotfields-object-excel.md)** object) that are currently shown as column fields. Read-only.


## Syntax

 _expression_ . **ColumnFields**( **_Index_** )

 _expression_ A variable that represents a **PivotTable** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|The field name or number (can be an array to specify more than one field).|

## Example

This example adds the field names of the PivotTable report columns to a list on a new worksheet.


```vb
Set nwSheet = Worksheets.Add 
nwSheet.Activate 
Set pvtTable = Worksheets("Sheet2").Range("A1").PivotTable 
rw = 0 
For Each pvtField In pvtTable.ColumnFields 
 rw = rw + 1 
 nwSheet.Cells(rw, 1).Value = pvtField.Name 
Next pvtField
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

