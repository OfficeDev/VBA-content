---
title: Outline.ShowLevels Method (Excel)
keywords: vbaxl10.chm455074
f1_keywords:
- vbaxl10.chm455074
ms.prod: excel
api_name:
- Excel.Outline.ShowLevels
ms.assetid: 2ebeb135-bbb9-aac1-57d7-02a141aa3ddb
ms.date: 06/08/2017
---


# Outline.ShowLevels Method (Excel)

Displays the specified number of row and/or column levels of an outline.


## Syntax

 _expression_ . **ShowLevels**( **_RowLevels_** , **_ColumnLevels_** )

 _expression_ A variable that represents an **Outline** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _RowLevels_|Optional| **Variant**|Specifies the number of row levels of an outline to display. If the outline has fewer levels than the number specified, Microsoft Excel displays all the levels. If this argument is 0 (zero) or is omitted, no action is taken on rows.|
| _ColumnLevels_|Optional| **Variant**|Specifies the number of column levels of an outline to display. If the outline has fewer levels than the number specified, Microsoft Excel displays all the levels. If this argument is 0 (zero) or is omitted, no action is taken on columns.|

### Return Value

Variant


## Remarks

You must specify at least one argument.


## Example

This example displays row levels one through three and column level one of the outline on Sheet1.


```vb
Worksheets("Sheet1").Outline _ 
 .ShowLevels rowLevels:=3, columnLevels:=1
```


## See also


#### Concepts


[Outline Object](outline-object-excel.md)

