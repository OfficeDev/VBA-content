---
title: SmartArtNodes Object (Office)
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.SmartArtNodes
ms.assetid: 4c35e5a4-15a1-dd6d-85a2-eb30cbaa3093
---


# SmartArtNodes Object (Office)

Represents a collection of nodes within a Smart Art diagram. 


## Remarks

These nodes correspond directly to semantic elements contained within the data model of the graphic.


## Example

The following code returns the number of nodes in the Smart Art diagram.


```vb
ActivePresentation.Slides(1).Shapes(1).SmartArtNodes.Count
```


## See also


#### Concepts


[Object Model Reference](../../Office-Shared-VBA/articles/reference-object-library-reference-for-office.md)

