---
title: Columns.Add Method (Publisher)
keywords: vbapb10.chm5046276
f1_keywords:
- vbapb10.chm5046276
ms.prod: publisher
api_name:
- Publisher.Columns.Add
ms.assetid: b3dfb892-6bda-d2c4-11f7-9bd29bf257aa
ms.date: 06/08/2017
---


# Columns.Add Method (Publisher)

Adds a new  **Column** object to the specified **Columns** collection and returns the new **Column** object.


## Syntax

 _expression_. **Add**( **_BeforeColumn_**)

 _expression_A variable that represents a  **Columns** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|BeforeColumn|Optional| **Long**|The number of the column before which to insert the new column. If this argument is omitted, the new column is added after the existing columns. An error occurs if the value of this argument does not correspond to an existing column in the table.|

### Return Value

Column


## Example

The following example adds a column before column three in the specified table.


```vb
Dim colNew As Column 
 
Set colNew = ActiveDocument.Pages(1).Shapes(1) _ 
 .Table.Columns.Add(BeforeColumn:=3)
```


