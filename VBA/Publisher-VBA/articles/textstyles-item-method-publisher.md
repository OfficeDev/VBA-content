---
title: TextStyles.Item Method (Publisher)
keywords: vbapb10.chm5898240
f1_keywords:
- vbapb10.chm5898240
ms.prod: publisher
api_name:
- Publisher.TextStyles.Item
ms.assetid: 14d1871f-c2cb-31af-e22d-10b3cf59b6fc
ms.date: 06/08/2017
---


# TextStyles.Item Method (Publisher)

Returns an individual object in a specified collection.


## Syntax

 _expression_. **Item**( **_Index_**)

 _expression_A variable that represents a  **TextStyles** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Index|Required| **Variant**|The number or name of the field or list box item to return.|

### Return Value

TextStyle


## Example

This example returns the "Normal" text style from the active publication.


```vb
Dim txtStyle As TextStyle 
 
Set txtStyle = ActiveDocument.TextStyles.Item(Index:="Normal")
```


