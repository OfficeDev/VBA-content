---
title: Sort.SetRange Method (Excel)
keywords: vbaxl10.chm847079
f1_keywords:
- vbaxl10.chm847079
ms.prod: excel
api_name:
- Excel.Sort.SetRange
ms.assetid: 12a68fb7-379d-f9fa-d464-a6d5fe1e6f9b
ms.date: 06/08/2017
---


# Sort.SetRange Method (Excel)

Sets the range over which the sort occurs.


## Syntax

 _expression_ . **SetRange**( **_Rng_** )

 _expression_ A variable that represents a **Sort** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Rng_|Required| **Range**|Specifies the range over which the sort represented by the  **Sort** object occurs.|

 **Note**   **SetRange** can only be used when applying a sort to a sheet range, and cannot be used if the range is within a table.


## See also


#### Concepts


[Sort Object](sort-object-excel.md)

