---
title: CustomViews.Add Method (Excel)
keywords: vbaxl10.chm506075
f1_keywords:
- vbaxl10.chm506075
ms.prod: excel
api_name:
- Excel.CustomViews.Add
ms.assetid: 134d9969-048b-6a53-4f2c-cc83589c5a70
ms.date: 06/08/2017
---


# CustomViews.Add Method (Excel)

Creates a new custom view.


## Syntax

 _expression_ . **Add**( **_ViewName_** , **_PrintSettings_** , **_RowColSettings_** )

 _expression_ A variable that represents a **CustomViews** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ViewName_|Required| **String**|The name of the new view.|
| _PrintSettings_|Optional| **Variant**| **True** to include print settings in the custom view.|
| _RowColSettings_|Optional| **Variant**| **True** to include settings for hidden rows and columns (including filter information) in the custom view.|

### Return Value

A  **[CustomView](customview-object-excel.md)** object that represents the new custom view.


## Example

This example creates a new custom view named "Summary" in the active workbook.


```vb
ActiveWorkbook.CustomViews.Add "Summary", True, True
```


## See also


#### Concepts


[CustomViews Object](customviews-object-excel.md)

