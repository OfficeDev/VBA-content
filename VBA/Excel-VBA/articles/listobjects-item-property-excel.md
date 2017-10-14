---
title: ListObjects.Item Property (Excel)
keywords: vbaxl10.chm732076
f1_keywords:
- vbaxl10.chm732076
ms.prod: excel
api_name:
- Excel.ListObjects.Item
ms.assetid: 39f00da9-170d-e62b-4beb-38e06a8ba533
ms.date: 06/08/2017
---


# ListObjects.Item Property (Excel)

Returns a single object from a collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **ListObjects** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or index number of the object.|

## Example

The following example displays the name of the default list object on Sheet1 of the active workbook.


```vb
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set oListObj = wrksht.ListObjects.Item(1).Name
```


## See also


#### Concepts


[ListObjects Object](listobjects-object-excel.md)

