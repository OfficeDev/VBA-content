---
title: CellRange.Item Method (Publisher)
keywords: vbapb10.chm5177344
f1_keywords:
- vbapb10.chm5177344
ms.prod: publisher
api_name:
- Publisher.CellRange.Item
ms.assetid: 8f1fe143-e00c-7112-45dd-52158153cf28
ms.date: 06/08/2017
---


# CellRange.Item Method (Publisher)

Returns an individual  **Cell** object in the specified **CellRange** collection.


## Syntax

 _expression_. **Item**( **_Index_**)

 _expression_A variable that represents a  **CellRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Index|Required| **Long**|The number of the object to return.|

### Return Value

Cell


## Example

This example returns the first cell from a  **CellRange** collection.


```vb
Dim cllTemp As Cell 
 
Set cllTemp = ActiveDocument.Pages(Index:=1).Shapes(1).Table.Cells.Item(Index:=1)
```


