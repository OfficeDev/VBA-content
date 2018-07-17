---
title: PivotTable.VisibleFields Property (Excel)
keywords: vbaxl10.chm235101
f1_keywords:
- vbaxl10.chm235101
ms.prod: excel
api_name:
- Excel.PivotTable.VisibleFields
ms.assetid: 01d5e76d-e109-905d-1743-1fbacd85e7a6
ms.date: 06/08/2017
---


# PivotTable.VisibleFields Property (Excel)

Returns an object that represents either a single field in a PivotTable report (a  **[PivotField](pivotfield-object-excel.md)** object) or a collection of all the visible fields (a **[PivotFields](pivotfields-object-excel.md)** object). Visible fields are shown as row, column, page or data fields. Read-only.


## Syntax

 _expression_ . **VisibleFields**( **_Index_** )

 _expression_ A variable that represents a **PivotTable** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|The name or number of the field to be returned (can be an array to specify more than one field).|

## Remarks

For OLAP data sources, there are no hidden fields, and this property returns all the fields in the PivotTable cache.


## Example

This example adds the visible field names to a list on a new worksheet.


```vb
Set nwSheet = Worksheets.Add 
nwSheet.Activate 
Set pvtTable = Worksheets("Sheet2").Range("A1").PivotTable 
rw = 0 
For Each pvtField In pvtTable.VisibleFields 
 rw = rw + 1 
 nwSheet.Cells(rw, 1).Value = pvtField.Name 
Next pvtField
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

