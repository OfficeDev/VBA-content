---
title: PivotTable.PageFields Property (Excel)
keywords: vbaxl10.chm235086
f1_keywords:
- vbaxl10.chm235086
ms.prod: excel
api_name:
- Excel.PivotTable.PageFields
ms.assetid: eff7a772-0472-41ec-412f-9a56f0a0de16
ms.date: 06/08/2017
---


# PivotTable.PageFields Property (Excel)

Returns an object that represents either a single PivotTable field (a  **[PivotField](pivotfield-object-excel.md)** object) or a collection of all the fields (a **[PivotFields](pivotfields-object-excel.md)** object) that are currently showing as page fields. Read-only.


## Syntax

 _expression_ . **PageFields**( **_Index_** )

 _expression_ A variable that represents a **PivotTable** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|The name or number of the field to be returned (can be an array to specify more than one field).|

## Remarks

A hierarchy can contain only one page field.

For a PivotTable report based on a PivotTable cache, the collection of PivotTable fields that?s returned reflects what?s currently in the cache.


## Example

This example adds the page field names to a list on a new worksheet.


```vb
Set nwSheet = Worksheets.Add 
nwSheet.Activate 
Set pvtTable = Worksheets("Sheet2").Range("A1").PivotTable 
rw = 0 
For Each pvtField In pvtTable.PageFields 
 rw = rw + 1 
 nwSheet.Cells(rw, 1).Value = pvtField.Name 
Next pvtField 

```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

