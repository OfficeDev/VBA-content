---
title: Shape.Apply Method (Publisher)
keywords: vbapb10.chm2228240
f1_keywords:
- vbapb10.chm2228240
ms.prod: publisher
api_name:
- Publisher.Shape.Apply
ms.assetid: 711c72b6-3618-be0b-fb72-9f68fdbcc4a8
ms.date: 06/08/2017
---


# Shape.Apply Method (Publisher)

Applies formatting copied from another shape or shape range using the  **[PickUp](shape-pickup-method-publisher.md)** method.


## Syntax

 _expression_. **Apply**

 _expression_A variable that represents a  **Shape** object.


### Return Value

Nothing


## Remarks

If you do not first use the  **PickUp** method to copy the formatting from another shape, an error occurs.


## Example

The following example copies the formatting from the first shape of the active publication to the second shape of the active publication.


```vb
With ActiveDocument.Pages(1) 
 .Shapes(1).PickUp 
 .Shapes(2).Apply 
End With 

```


