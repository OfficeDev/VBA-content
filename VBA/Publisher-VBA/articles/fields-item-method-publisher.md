---
title: Fields.Item Method (Publisher)
keywords: vbapb10.chm6029312
f1_keywords:
- vbapb10.chm6029312
ms.prod: publisher
api_name:
- Publisher.Fields.Item
ms.assetid: 95783e5a-2c82-235e-75a4-5ac15938718e
ms.date: 06/08/2017
---


# Fields.Item Method (Publisher)

Returns an individual object in a specified collection.


## Syntax

 _expression_. **Item**( **_Index_**)

 _expression_A variable that represents a  **Fields** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Index|Required| **Long**|The number of the object to return.|

### Return Value

Field


## Example

This example returns the first field from a  **Fields** object.


```vb
Dim fldTemp As Field 
 
Set fldTemp = ActiveDocument.Pages(Index:=1) _ 
 .Shapes(1).TextFrame.TextRange.Fields.Item(Index:=1)
```


