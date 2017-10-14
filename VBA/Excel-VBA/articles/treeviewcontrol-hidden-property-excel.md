---
title: TreeviewControl.Hidden Property (Excel)
keywords: vbaxl10.chm666073
f1_keywords:
- vbaxl10.chm666073
ms.prod: excel
api_name:
- Excel.TreeviewControl.Hidden
ms.assetid: 134a3b6b-492b-6813-cd40-ce1ff3b52c6c
ms.date: 06/08/2017
---


# TreeviewControl.Hidden Property (Excel)

Returns or sets a Variant value that represents the hidden status of the cube field members in the hierarchical member selection control of a cube field.


## Syntax

 _expression_ . **Hidden**

 _expression_ A variable that represents a **TreeviewControl** object.


## Remarks

Don't confuse this property with the  **[FormulaHidden](range-formulahidden-property-excel.md)** property.

The  **Hidden** property returns or sets an array. Each element of the array corresponds to a level of the cube field that is hidden. The maximum number of elements is the number of levels in the cube field. Each element of the array is an array of type **String** , containing unique member names that are hidden at the corresponding level of the control. See the **[DrilledDown](pivotitem-drilleddown-property-excel.md)** property of the **[PivotItem](pivotitem-object-excel.md)** object to determine when members are visible (expanded) in the control.


## Example

This example hides the second level member [state].[states].[CA].[Covelo] of the first cube field in the first PivotTable report.


```vb
ActiveSheet.PivotTables("PivotTable1").CubeFields(1) _ 
 .TreeviewControl.Hidden = _ 
 Array(Array(""), Array(""), _ 
 Array("[state].[states].[CA].[Covelo]"))
```


## See also


#### Concepts


[TreeviewControl Object](treeviewcontrol-object-excel.md)

