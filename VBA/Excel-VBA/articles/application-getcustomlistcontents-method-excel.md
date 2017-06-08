---
title: Application.GetCustomListContents Method (Excel)
keywords: vbaxl10.chm133140
f1_keywords:
- vbaxl10.chm133140
ms.prod: excel
api_name:
- Excel.Application.GetCustomListContents
ms.assetid: 3adafb35-f7d0-0233-ff7c-c31d5e48f574
ms.date: 06/08/2017
---


# Application.GetCustomListContents Method (Excel)

Returns a custom list (an array of strings).


## Syntax

 _expression_ . **GetCustomListContents**( **_ListNum_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ListNum_|Required| **Long**|The list number.|

### Return Value

Variant


## Example

This example writes the elements of the first custom list in column one on Sheet1.


```vb
listArray = Application.GetCustomListContents(1) 
For i = LBound(listArray, 1) To UBound(listArray, 1) 
 Worksheets("sheet1").Cells(i, 1).Value = listArray(i) 
Next i
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

