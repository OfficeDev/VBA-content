---
title: LegendEntries.Item Method (Excel)
keywords: vbaxl10.chm588075
f1_keywords:
- vbaxl10.chm588075
ms.prod: excel
api_name:
- Excel.LegendEntries.Item
ms.assetid: 8f7250b8-1c52-3e8a-4b09-906e917fdcac
ms.date: 06/08/2017
---


# LegendEntries.Item Method (Excel)

Returns a single object from a collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **LegendEntries** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The index number for the object.|

### Return Value

A  **[LegendEntry](legendentry-object-excel.md)** object contained by the collection.


## Example

This example changes the font for the text of the legend entry at the top of the legend (this is usually the legend for series one) in embedded chart one on Sheet1.


```vb
Worksheets("sheet1").ChartObjects(1).Chart _ 
 .Legend.LegendEntries.Item(1).Font.Italic = True
```


## See also


#### Concepts


[LegendEntries Object](legendentries-object-excel.md)

