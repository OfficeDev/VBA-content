---
title: Field.Unlink Method (Publisher)
keywords: vbapb10.chm6094857
f1_keywords:
- vbapb10.chm6094857
ms.prod: publisher
api_name:
- Publisher.Field.Unlink
ms.assetid: 4dfe5c29-eb1e-b071-fd86-6ee222455c4e
ms.date: 06/08/2017
---


# Field.Unlink Method (Publisher)

Replaces the specified field or  **[Fields](fields-object-publisher.md)** collection with with their most recent results.


## Syntax

 _expression_. **Unlink**

 _expression_A variable that represents a  **Field** object.


## Remarks

When you unlink a field, its current result is converted to text or a graphic and can no longer be updated automatically.


## Example

This example unlinks the first field in shape one on the first page of the active publication.


```vb
ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.Fields(1).Unlink
```


