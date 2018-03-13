---
title: Borders.Item Property (Excel)
keywords: vbaxl10.chm181076
f1_keywords:
- vbaxl10.chm181076
ms.prod: excel
api_name:
- Excel.Borders.Item
ms.assetid: 19184379-d551-396e-8cb6-ff240e3c85fa
ms.date: 06/08/2017
---


# Borders.Item Property (Excel)

Returns a  **[Border](border-object-excel.md)** object that represents one of the borders of either a range of cells or a style.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **Borders** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **XlBordersIndex**|One of the constants of  **XlBordersIndex** .|

## Remarks





| <strong>XlBordersIndex</strong> can be one of these <strong>XlBordersIndex</strong> constants.|
| 
<strong>xlDiagonalDown</strong>|
| 
<strong>xlDiagonalUp</strong>|
| 
<strong>xlEdgeBottom</strong>|
| 
<strong>xlEdgeLeft</strong>|
| 
<strong>xlEdgeRight</strong>|
| 
<strong>xlEdgeTop</strong>|
| 
<strong>xlInsideHorizontal</strong>|
| 
<strong>xlInsideVertical</strong>|

## Example

This following example sets the color of the bottom border of cells A1:G1.


```vb
Worksheets("Sheet1").Range("a1:g1"). _ 
 Borders.Item(xlEdgeBottom).Color = RGB(255, 0, 0)
```


## See also


#### Concepts


[Borders Collection](borders-object-excel.md)

