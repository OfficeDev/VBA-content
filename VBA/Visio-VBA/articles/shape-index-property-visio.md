---
title: Shape.Index Property (Visio)
keywords: vis_sdr.chm11213695
f1_keywords:
- vis_sdr.chm11213695
ms.prod: visio
api_name:
- Visio.Shape.Index
ms.assetid: 7fb67e8b-76a7-c2ac-e7aa-89635ca7622c
ms.date: 06/08/2017
---


# Shape.Index Property (Visio)

Gets the ordinal position of a  **Shape** object in the **Shapes** collection. Read-only.


## Syntax

 _expression_ . **Index**

 _expression_ A variable that represents a **Shape** object.


### Return Value

Long


## Remarks

Most collections are indexed starting with 1 rather than zero (0), so the index of the first element is 1, the index of the second element is 2, and so forth. The index of the last element in a collection is the same as the value of that collection's  **Count** property. You can iterate through a collection by using these index values. Adding objects to or deleting objects from a collection can change the index values of other objects in the collection.


