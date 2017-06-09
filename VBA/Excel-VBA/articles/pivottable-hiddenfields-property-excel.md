---
title: PivotTable.HiddenFields Property (Excel)
keywords: vbaxl10.chm235083
f1_keywords:
- vbaxl10.chm235083
ms.prod: excel
api_name:
- Excel.PivotTable.HiddenFields
ms.assetid: f59f471f-5ce9-fa81-ab37-91eb78666870
ms.date: 06/08/2017
---


# PivotTable.HiddenFields Property (Excel)

Returns an object that represents either a single PivotTable field (a  **[PivotField](pivotfield-object-excel.md)** object) or a collection of all the fields (a **[PivotFields](pivotfields-object-excel.md)** object) that are currently not shown as row, column, page, or data fields. Read-only.


## Syntax

 _expression_ . **HiddenFields**( **_Index_** )

 _expression_ A variable that represents a **PivotTable** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|The name or number of the field to be returned (can be an array to specify more than one field).|

## Remarks

For OLAP data sources, this property always returns an empty collection.


## Example

This example adds the hidden field names to a list on a new worksheet.


```vb
Set nwSheet = Worksheets.Add 
nwSheet.Activate 
Set pvtTable = Worksheets("Sheet2").Range("A1").PivotTable 
rw = 0 
For Each pvtField In pvtTable.HiddenFields 
 rw = rw + 1 
 nwSheet.Cells(rw, 1).Value = pvtField.Name 
Next pvtField
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

