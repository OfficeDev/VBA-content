---
title: Columns.Item Method (Publisher)
keywords: vbapb10.chm5046272
f1_keywords:
- vbapb10.chm5046272
ms.prod: publisher
api_name:
- Publisher.Columns.Item
ms.assetid: c16df25c-ea8d-c04e-bccd-7e642bb7198a
ms.date: 06/08/2017
---


# Columns.Item Method (Publisher)

Returns an individual  **Column** object in the specified **Columns** collection.


## Syntax

 _expression_. **Item**( **_Index_**)

 _expression_A variable that represents a  **Columns** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Index|Required| **Long**|The number of the object to return.|

### Return Value

Column


## Example

This example returns the first column from a  **Columns** collection.


```vb
Dim colTemp As Column 
 
Set colTemp = ActiveDocument.Pages(Index:=1) _ 
 .Shapes(1).Table.Columns.Item(Index:=1)
```


