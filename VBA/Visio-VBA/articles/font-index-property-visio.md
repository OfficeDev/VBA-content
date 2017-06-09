---
title: Font.Index Property (Visio)
keywords: vis_sdr.chm12013695
f1_keywords:
- vis_sdr.chm12013695
ms.prod: visio
api_name:
- Visio.Font.Index
ms.assetid: 3abb359d-0a40-1f2d-5d00-041cc6d39385
ms.date: 06/08/2017
---


# Font.Index Property (Visio)

Gets the ordinal position of a  **Font** object in the **Fonts** collection. Read-only.


## Syntax

 _expression_ . **Index**

 _expression_ A variable that represents a **Font** object.


### Return Value

Integer


## Remarks

Most collections are indexed starting with 1 rather than zero (0), so the index of the first element is 1, the index of the second element is 2, and so forth. The index of the last element in a collection is the same as the value of that collection's  **Count** property. You can iterate through a collection by using these index values. Adding objects to or deleting objects from a collection can change the index values of other objects in the collection.


