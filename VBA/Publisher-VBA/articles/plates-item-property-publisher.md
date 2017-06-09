---
title: Plates.Item Property (Publisher)
keywords: vbapb10.chm2818048
f1_keywords:
- vbapb10.chm2818048
ms.prod: publisher
api_name:
- Publisher.Plates.Item
ms.assetid: 7563df76-56c3-d613-7314-846fe28a995d
ms.date: 06/08/2017
---


# Plates.Item Property (Publisher)

Returns an individual object from a specified collection. Read-only.


## Syntax

 _expression_. **Item**( **_Index_**)

 _expression_A variable that represents a  **Plates** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Index|Required| **Long**|The number of the object to return.|

## Example

This example displays the name of the first color plate in the active publication.


```vb
MsgBox "Name of first color plate: " _ 
 &; ActiveDocument.Plates.Item(1).Name
```


