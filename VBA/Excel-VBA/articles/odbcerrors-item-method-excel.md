---
title: ODBCErrors.Item Method (Excel)
keywords: vbaxl10.chm529074
f1_keywords:
- vbaxl10.chm529074
ms.prod: excel
api_name:
- Excel.ODBCErrors.Item
ms.assetid: 694a0e7e-f6c0-8721-792b-8e82e6a8e5c1
ms.date: 06/08/2017
---


# ODBCErrors.Item Method (Excel)

Returns a single object from a collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents an **ODBCErrors** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|The index number for the object.|

### Return Value

An  **[ODBCError](odbcerror-object-excel.md)** object contained by the collection.


## Example

This example displays an ODBC error.


```vb
Set er = Application.ODBCErrors.Item(1) 
MsgBox "The following error occurred:" &; 
 er.ErrorString &; " : " &; er.SqlState
```


## See also


#### Concepts


[ODBCErrors Object](odbcerrors-object-excel.md)

