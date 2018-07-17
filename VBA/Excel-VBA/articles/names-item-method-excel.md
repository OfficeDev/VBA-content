---
title: Names.Item Method (Excel)
keywords: vbaxl10.chm488074
f1_keywords:
- vbaxl10.chm488074
ms.prod: excel
api_name:
- Excel.Names.Item
ms.assetid: 01d138f1-a2a8-8c39-98f0-b953c4b3b5ba
ms.date: 06/08/2017
---


# Names.Item Method (Excel)

Returns a single  **[Name](name-object-excel.md)** object from a **Names** collection.


## Syntax

 _expression_ . **Item**( **_Index_** , **_IndexLocal_** , **_RefersTo_** )

 _expression_ A variable that represents a **Names** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|The name or number of the defined name to be returned.|
| _IndexLocal_|Optional| **Variant**|The name of the defined name, in the language of the user. No names will be translated if you use this argument.|
| _RefersTo_|Optional| **Variant**|What the name refers to. You use this argument to identify a name by what it refers to.|

### Return Value

A  **Name** object contained by the collection.


## Remarks

You must specify one, and only one, of these three arguments.


## Example

This example deletes the name mySortRange from the active workbook.


```vb
ActiveWorkbook.Names.Item("mySortRange").Delete
```


## See also


#### Concepts


[Names Object](names-object-excel.md)

