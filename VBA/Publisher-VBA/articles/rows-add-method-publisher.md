---
title: Rows.Add Method (Publisher)
keywords: vbapb10.chm4915204
f1_keywords:
- vbapb10.chm4915204
ms.prod: publisher
api_name:
- Publisher.Rows.Add
ms.assetid: 34d72709-92f7-ddc6-5be6-e74693466e61
ms.date: 06/08/2017
---


# Rows.Add Method (Publisher)

Adds a new  **Row** object to the specified **Rows** collection and returns the new **Row** object.


## Syntax

 _expression_. **Add**( **_BeforeRow_**)

 _expression_A variable that represents a  **Rows** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|BeforeRow|Optional| **Long**|The number of the row before which to insert the new row. If this argument is omitted, the new row is added after the existing rows. An error occurs if the value of this argument does not correspond to an existing row in the table.|

### Return Value

Row


## Example

The following example adds a row before row three in the specified table.


```vb
Dim rowNew As Row 
 
Set rowNew = ActiveDocument.Pages(1).Shapes(1) _ 
 .Table.Rows.Add(BeforeRow:=3)
```


