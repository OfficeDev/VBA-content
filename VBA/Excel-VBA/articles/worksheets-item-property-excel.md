---
title: Worksheets.Item Property (Excel)
keywords: vbaxl10.chm470078
f1_keywords:
- vbaxl10.chm470078
ms.prod: excel
api_name:
- Excel.Worksheets.Item
ms.assetid: 66099ca2-54a0-f8ae-58ab-07791bbf5e7c
ms.date: 06/08/2017
---


# Worksheets.Item Property (Excel)

Returns a single object from a collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **Worksheets** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or index number of the object.|

## Remarks

For more information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/f8a36459-f9dd-9f4c-ef7a-b188173434d5%28Office.15%29.aspx).


## Example

 **Item** is the default member for a collection. For example, the following two lines of code are equivalent.


```vb
ActiveWorkbook.Worksheets.Item(1) 
ActiveWorkbook.Worksheets(1)
```


## See also


#### Concepts


[Worksheets Object](worksheets-object-excel.md)

