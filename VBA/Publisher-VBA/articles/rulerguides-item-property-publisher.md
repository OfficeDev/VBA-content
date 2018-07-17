---
title: RulerGuides.Item Property (Publisher)
keywords: vbapb10.chm720896
f1_keywords:
- vbapb10.chm720896
ms.prod: publisher
api_name:
- Publisher.RulerGuides.Item
ms.assetid: e0c49279-4fd4-fe61-636c-c29399fdc404
ms.date: 06/08/2017
---


# RulerGuides.Item Property (Publisher)

Returns an individual object from a specified collection. Read-only.


## Syntax

 _expression_. **Item**( **_Index_**)

 _expression_A variable that represents a  **RulerGuides** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Item|Required| **Long**|The number of the object to return.|

## Example

This example sets the position of the first ruler guide to 3 inches from the edge of the publication.


```vb
ActiveDocument.Pages(1).RulerGuides _ 
 .Item(1).Position = InchesToPoints(3)
```


