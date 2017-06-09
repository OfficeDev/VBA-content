---
title: PivotTable.ErrorString Property (Excel)
keywords: vbaxl10.chm235109
f1_keywords:
- vbaxl10.chm235109
ms.prod: excel
api_name:
- Excel.PivotTable.ErrorString
ms.assetid: 7f00d151-9f92-a3b3-c95f-60c0600cf594
ms.date: 06/08/2017
---


# PivotTable.ErrorString Property (Excel)

Returns or sets a  **String** value that represents the string displayed in cells that contain errors when the **[DisplayErrorString](pivottable-displayerrorstring-property-excel.md)** property is **True** .


## Syntax

 _expression_ . **ErrorString**

 _expression_ A variable that represents a **PivotTable** object.


## Remarks

The default value for this property is an empty string ("").


## Example

This example displays a hyphen in cells in the specified PivotTable report that contain errors.


```vb
With Worksheets(1).PivotTables("Pivot1") 
 .ErrorString = "-" 
 .DisplayErrorString = True 
End With
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

