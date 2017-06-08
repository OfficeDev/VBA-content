---
title: SlicerPivotTables.RemovePivotTable Method (Excel)
keywords: vbaxl10.chm911078
f1_keywords:
- vbaxl10.chm911078
ms.prod: excel
api_name:
- Excel.SlicerPivotTables.RemovePivotTable
ms.assetid: ebc4cc53-c406-3ae4-06e7-094a1ba32af2
ms.date: 06/08/2017
---


# SlicerPivotTables.RemovePivotTable Method (Excel)

Removes a reference to a PivotTable from the  **[SlicerPivotTables](slicerpivottables-object-excel.md)** collection.


## Syntax

 _expression_ . **RemovePivotTable**( **_PivotTable_** )

 _expression_ A variable that represents a **SlicerPivotTables** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _PivotTable_|Required| **Variant**|A  **[PivotTable](pivottable-object-excel.md)** object that represents the PivotTable to remove, or the name or index of the PivotTable in the collection.|

### Return Value

Nothing


## Remarks

When a PivotTable is removed from the  **SlicerPivotTables** collection, is no longer filtered by its parent **[SlicerCache](slicercache-object-excel.md)** and the slicers associated with it.


## Example

The following code example removes PivotTable1 from the slicer cache associated with the Customer slicer.


```vb
Dim pvts As SlicerPivotTables 
Set pvts = ActiveWorkbook.SlicerCaches("Slicer_Customer").PivotTables 
pvts.RemovePivotTable("PivotTable1")
```


## See also


#### Concepts


[SlicerPivotTables Object](slicerpivottables-object-excel.md)

