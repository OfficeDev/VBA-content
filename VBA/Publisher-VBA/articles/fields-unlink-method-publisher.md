---
title: Fields.Unlink Method (Publisher)
keywords: vbapb10.chm6029316
f1_keywords:
- vbapb10.chm6029316
ms.prod: publisher
api_name:
- Publisher.Fields.Unlink
ms.assetid: 7a40909f-5fc1-84ef-6679-969a98a8a668
ms.date: 06/08/2017
---


# Fields.Unlink Method (Publisher)

Replaces the specified field or  **[Fields](fields-object-publisher.md)** collection with with their most recent results.


## Syntax

 _expression_. **Unlink**

 _expression_A variable that represents a  **Fields** object.


### Return Value

Nothing


## Remarks

When you unlink a field, its current result is converted to text or a graphic and can no longer be updated automatically.


## Example

This example unlinks the first field in shape one on the first page of the active publication.


```vb
ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.Fields(1).Unlink
```


