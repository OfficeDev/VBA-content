---
title: PivotField.ParentItems Property (Excel)
keywords: vbaxl10.chm240090
f1_keywords:
- vbaxl10.chm240090
ms.prod: excel
api_name:
- Excel.PivotField.ParentItems
ms.assetid: 361db264-aa5a-9547-5405-41203fe3df0a
ms.date: 06/08/2017
---


# PivotField.ParentItems Property (Excel)

Returns an object that represents either a single PivotTable item (a  **[PivotItem](pivotitem-object-excel.md)** object) or a collection of all the items (a **[PivotItems](pivotitems-object-excel.md)** object) that are group parents in the specified field. The specified field must be a group parent of another field. Read-only.


## Syntax

 _expression_ . **ParentItems**( **_Index_** )

 _expression_ A variable that represents a **PivotField** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|The number or name of the item to be returned (can be an array to specify more than one item).|

## Remarks

This property isn?t available for OLAP data sources.


## Example

This example creates a list containing the names of all the items that are group parents in the field named "product".


```vb
Set nwSheet = Worksheets.Add 
nwSheet.Activate 
Set pvtTable = Worksheets("Sheet2").Range("A1").PivotTable 
rw = 0 
For Each pvtItem In pvtTable.PivotFields("product").ParentItems 
 rw = rw + 1 
 nwSheet.Cells(rw, 1).Value = pvtItem.Name 
Next pvtItem
```


## See also


#### Concepts


[PivotField Object](pivotfield-object-excel.md)

