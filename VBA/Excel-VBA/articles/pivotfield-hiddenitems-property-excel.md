---
title: PivotField.HiddenItems Property (Excel)
keywords: vbaxl10.chm240083
f1_keywords:
- vbaxl10.chm240083
ms.prod: excel
api_name:
- Excel.PivotField.HiddenItems
ms.assetid: ec30c18e-c030-23b8-2ea8-7ed7bfbd3312
ms.date: 06/08/2017
---


# PivotField.HiddenItems Property (Excel)

Returns an object that represents either a single hidden PivotTable item (a  **[PivotItem](pivotitem-object-excel.md)** object) or a collection of all the hidden items (a **[PivotItems](pivotitems-object-excel.md)** object) in the specified field. Read-only.


## Syntax

 _expression_ . **HiddenItems**( **_Index_** )

 _expression_ A variable that represents a **PivotField** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|The number or name of the item to be returned (can be an array to specify more than one item).|

## Remarks

For OLAP data sources, this property always returns an empty collection.


## Example

This example adds the names of all the hidden items in the field named "product" to a list on a new worksheet.


```vb
Set nwSheet = Worksheets.Add 
nwSheet.Activate 
Set pvtTable = Worksheets("Sheet2").Range("A1").PivotTable 
rw = 0 
For Each pvtItem In pvtTable.PivotFields("product").HiddenItems 
 rw = rw + 1 
 nwSheet.Cells(rw, 1).Value = pvtItem.Name 
Next pvtItem
```


## See also


#### Concepts


[PivotField Object](pivotfield-object-excel.md)

