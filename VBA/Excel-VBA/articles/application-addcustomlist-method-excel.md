---
title: Application.AddCustomList Method (Excel)
keywords: vbaxl10.chm133076
f1_keywords:
- vbaxl10.chm133076
ms.prod: excel
api_name:
- Excel.Application.AddCustomList
ms.assetid: 31518c3c-78ce-f9e9-9572-a1338aa6d2e7
ms.date: 06/08/2017
---


# Application.AddCustomList Method (Excel)

Adds a custom list for custom autofill and/or custom sort.


## Syntax

 _expression_ . **AddCustomList**( **_ListArray_** , **_ByRow_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ListArray_|Required| **Variant**|Specifies the source data, as either an array of strings or a  **Range** object.|
| _ByRow_|Optional| **Variant**|Only used if  _ListArray_ is a **Range** object. **True** to create a custom list from each row in the range. **False** to create a custom list from each column in the range. If this argument is omitted and there are more rows than columns (or an equal number of rows and columns) in the range, Microsoft Excel creates a custom list from each column in the range. If this argument is omitted and there are more columns than rows in the range, Microsoft Excel creates a custom list from each row in the range.|

## Remarks

If the list you're trying to add already exists, this method throws a run-time error '1004'. Catch the error with an On Error statement.


## Example

This example adds an array of strings as a custom list.


```vb
On Error Resume Next  ' if the list already exists, do nothing
Application.AddCustomList Array("cogs", "sprockets", _ 
 "widgets", "gizmos")
On Error Goto 0       ' resume regular error handling
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

