---
title: Filters.Item Property (Excel)
keywords: vbaxl10.chm540075
f1_keywords:
- vbaxl10.chm540075
ms.prod: excel
api_name:
- Excel.Filters.Item
ms.assetid: a24c9aeb-b253-c11a-29dc-c4a2bba86e21
ms.date: 06/08/2017
---


# Filters.Item Property (Excel)

Returns a single object from a collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **Filters** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|The index number of the object.|

## Example

The following example sets a variable to the value of the  **On** property of the filter for the first column in the filtered range on the Crew worksheet.


```vb
Set w = Worksheets("Crew") 
If w.AutoFilterMode Then 
 filterIsOn = w.AutoFilter.Filters.Item(1).On 
End If
```


## See also


#### Concepts


[Filters Object](filters-object-excel.md)

