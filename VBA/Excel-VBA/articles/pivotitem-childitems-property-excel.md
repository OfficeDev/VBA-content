---
title: PivotItem.ChildItems Property (Excel)
keywords: vbaxl10.chm246074
f1_keywords:
- vbaxl10.chm246074
ms.prod: excel
api_name:
- Excel.PivotItem.ChildItems
ms.assetid: 5ae6936e-0ae7-284a-1733-86ba292e8a9c
ms.date: 06/08/2017
---


# PivotItem.ChildItems Property (Excel)

Returns an object that represents either a single PivotTable item (a  **[PivotItem](pivotitem-object-excel.md)** object) or a collection of all the items (a **[PivotItems](pivotitems-object-excel.md)** object) that are group children in the specified field, or children of the specified item. Read-only.


## Syntax

 _expression_ . **ChildItems**( **_Index_** )

 _expression_ A variable that represents a **PivotItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|The item name or number (can be an array to specify more than one item).|

## Remarks

This property is not available for OLAP data sources.


## Example

This example adds the names of all the child items of the item named "vegetables" to a list on a new worksheet.


```vb
Set nwSheet = Worksheets.Add 
nwSheet.Activate 
Set pvtTable = Worksheets("Sheet2").Range("A1").PivotTable 
rw = 0 
For Each pvtItem In _ 
 pvtTable.PivotFields("product") 
 .PivotItems("vegetables").ChildItems 
 rw = rw + 1 
 nwSheet.Cells(rw, 1).Value = pvtItem.Name 
Next pvtItem
```


## See also


#### Concepts


[PivotItem Object](pivotitem-object-excel.md)

