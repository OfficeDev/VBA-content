---
title: SmartArtNode.Shapes Property (Office)
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.SmartArtNode.Shapes
ms.assetid: c8a6dd3f-830e-342c-39c1-a86a54c475d4
---


# SmartArtNode.Shapes Property (Office)

Returns the shape range associated with this  **SmartArtNode** object. Read-only


## Syntax

 _expression_. **Shapes**

 _expression_ An expression that returns a **SmartArtNode** object.


## Remarks

To find the range, use the Count property.


## Example

The following code returns the shape range.


```vb
ActivePresentation.Slides(1).Shapes(1).SmartArtNodes.Item(1).Shapes.Count.
```


## See also


#### Concepts


[SmartArtNode Object](smartartnode-object-office.md)

