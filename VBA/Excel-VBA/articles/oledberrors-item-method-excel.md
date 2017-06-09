---
title: OLEDBErrors.Item Method (Excel)
keywords: vbaxl10.chm656074
f1_keywords:
- vbaxl10.chm656074
ms.prod: excel
api_name:
- Excel.OLEDBErrors.Item
ms.assetid: b5635b91-19ac-7915-ccb5-3bcb3d5d208a
ms.date: 06/08/2017
---


# OLEDBErrors.Item Method (Excel)

Returns a single object from a collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents an **OLEDBErrors** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|The index number for the object.|

### Return Value

An  **[OLEDBError](oledberror-object-excel.md)** object contained by the collection.


## Example

This example displays an OLE DB error.


```vb
Set objEr = Application.OLEDBErrors.Item(1) 
MsgBox "The following error occurred:" &; _ 
 objEr.ErrorString &; " : " &; objEr.SqlState
```


## See also


#### Concepts


[OLEDBErrors Object](oledberrors-object-excel.md)

