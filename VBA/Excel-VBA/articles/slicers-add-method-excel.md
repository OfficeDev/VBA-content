---
title: Slicers.Add Method (Excel)
keywords: vbaxl10.chm903077
f1_keywords:
- vbaxl10.chm903077
ms.prod: excel
api_name:
- Excel.Slicers.Add
ms.assetid: f2632dee-e8fb-440c-cad8-2dd2f7e37739
ms.date: 06/08/2017
---


# Slicers.Add Method (Excel)

Creates a new slicer and returns a  **[Slicer](slicer-object-excel.md)** object.


## Syntax

 _expression_ . **Add**( **_SlicerDestination_** , **_Level_** , **_Name_** , **_Caption_** , **_Top_** , **_Left_** , **_Width_** , **_Height_** )

 _expression_ A variable that represents a **Slicers** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SlicerDestination_|Required| **Variant**|A  **String** that specifies the name of the sheet, or a **Worksheet** object that represents the sheet, where the resulting slicer will be placed. The destination sheet must be in the workbook that contains the **Slicers** object specified by expression.|
| _Level_|Optional| **Variant**|For OLAP data sources, the ordinal or the Multidimensional Expression (MDX) name of the level on which the slicer creation is based. Not supported for non-OLAP data sources. |
| _Name_|Optional| **Variant**|The name of the slicer. Excel automatically generates a name if one is not specified. The name must be unique across all slicers within a workbook.|
| _Caption_|Optional| **Variant**|The caption of the slicer.|
| _Top_|Optional| **Variant**|The initial vertical position of the slicer, in points, relative to the upper-left corner of cell A1 on a worksheet.|
| _Left_|Optional| **Variant**|The initial horizontal position of the slicer, in points, relative to the upper-left corner of cell A1 on a worksheet.|
| _Width_|Optional| **Variant**|The initial width, in points, of the slicer control.|
| _Height_|Optional| **Variant**|The initial height, in points, of the slicer control.|

### Return Value

Slicer


## Example

This example adds a  **SlicerCache** object using the OLAP data source "AdventureWorks" and then adds a **Slicer** object to filter on the "Country" field.


```vb
Sub CreateNewSlicer() 
 ActiveWorkbook.SlicerCaches.Add("Adventure Works", _ 
 "[Customer].[Customer Geography]").Slicers.Add ActiveSheet, _ 
 "[Customer].[Customer Geography].[Country]", "Country 1", "Country", _ 
 252, 522, 144, 216) 
End Sub
```


## See also


#### Concepts


[Slicers Object](slicers-object-excel.md)

