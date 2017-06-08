---
title: Layer.Index Property (Visio)
keywords: vis_sdr.chm11813695
f1_keywords:
- vis_sdr.chm11813695
ms.prod: visio
api_name:
- Visio.Layer.Index
ms.assetid: 0a5d5637-7b93-f534-de65-c7fe230a6923
ms.date: 06/08/2017
---


# Layer.Index Property (Visio)

Gets the ordinal position of a  **Layer** object in the **Layers** collection. Read-only.


## Syntax

 _expression_ . **Index**

 _expression_ A variable that represents a **Layer** object.


### Return Value

Integer


## Remarks

Most collections are indexed starting with 1 rather than zero (0), so the index of the first element is 1, the index of the second element is 2, and so forth. The index of the last element in a collection is the same as the value of that collection's  **Count** property. You can iterate through a collection by using these index values. Adding objects to or deleting objects from a collection can change the index values of other objects in the collection.


