---
title: CustomViews.Item Method (Excel)
keywords: vbaxl10.chm506074
f1_keywords:
- vbaxl10.chm506074
ms.prod: excel
api_name:
- Excel.CustomViews.Item
ms.assetid: 542a3838-c499-5aa1-735e-5fe0c9c852a1
ms.date: 06/08/2017
---


# CustomViews.Item Method (Excel)

Returns a single object from a collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **CustomViews** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ViewName_|Required| **Variant**|The name or index number for the object.|

### Return Value

A  **[CustomView](customview-object-excel.md)** object contained by the collection.


## Example

This example includes print settings in the custom view named Current Inventory.


```vb
ThisWorkbook.CustomViews.Item("Current Inventory") _ 
 .PrintSettings = True
```


## See also


#### Concepts


[CustomViews Object](customviews-object-excel.md)

