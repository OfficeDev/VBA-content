---
title: Worksheets.Move Method (Excel)
keywords: vbaxl10.chm470079
f1_keywords:
- vbaxl10.chm470079
ms.prod: excel
api_name:
- Excel.Worksheets.Move
ms.assetid: e973d1d0-fd72-4e9e-e5b0-2b5d61eeed07
ms.date: 06/08/2017
---


# Worksheets.Move Method (Excel)

Moves the sheet to another location in the workbook.


## Syntax

 _expression_ . **Move**( **_Before_** , **_After_** )

 _expression_ A variable that represents a **Worksheets** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Before_|Optional| **Variant**|The sheet before which the moved sheet will be placed. You cannot specify  _Before_ if you specify _After_.|
| _After_|Optional| **Variant**| The sheet after which the moved sheet will be placed. You cannot specify _After_ if you specify _Before_.|

## Remarks

If you don't specify either  _Before_ or _After_, Microsoft Excel creates a new workbook that contains the moved sheet.


## Example

This example moves Sheet1 after Sheet3 in the active workbook.


```vb
Worksheets("Sheet1").Move _ 
 after:=Worksheets("Sheet3")
```


## See also


#### Concepts


[Worksheets Object](worksheets-object-excel.md)

