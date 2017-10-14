---
title: PivotField.PivotItems Method (Excel)
keywords: vbaxl10.chm240091
f1_keywords:
- vbaxl10.chm240091
ms.prod: excel
api_name:
- Excel.PivotField.PivotItems
ms.assetid: 5ec5fa1e-a080-2cbf-e4d4-b15d39e13ac5
ms.date: 06/08/2017
---


# PivotField.PivotItems Method (Excel)

Returns an object that represents either a single PivotTable item (a  **[PivotItem](pivotitem-object-excel.md)** object) or a collection of all the visible and hidden items (a **[PivotItems](pivotitems-object-excel.md)** object) in the specified field. Read-only.


## Syntax

 _expression_ . **PivotItems**( **_Index_** )

 _expression_ A variable that represents a **PivotField** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|The name or number of the item to be returned.|

### Return Value

Variant


## Remarks

For OLAP data sources, the collection is indexed by the unique name (the name returned by the  **[SourceName](pivotfield-sourcename-property-excel.md)** property), not by the display name.


## Example

This example adds the names of all items in the field named "product" to a list on a new worksheet.


```vb
Set nwSheet = Worksheets.Add 
nwSheet.Activate 
Set pvtTable = Worksheets("Sheet2").Range("A1").PivotTable 
rw = 0 
For Each pvtitem In pvtTable.PivotFields("product").PivotItems 
 rw = rw + 1 
 nwSheet.Cells(rw, 1).Value = pvtitem.Name 
Next
```


## See also


#### Concepts


[PivotField Object](pivotfield-object-excel.md)

