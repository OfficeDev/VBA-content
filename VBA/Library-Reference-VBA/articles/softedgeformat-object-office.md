---
title: SoftEdgeFormat Object (Office)
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.SoftEdgeFormat
ms.assetid: 9d9b34e1-03b5-9e56-b9ea-89c7ecce0370
---


# SoftEdgeFormat Object (Office)

Represents the soft edges effect in Office graphics.


## Remarks

The Soft Edge effect creates a mask around the edge of an object and blends the object with the transparent edge. The result is a faded or "feathered"edge.


## Example

This example sets the soft edge formatting for the text for the second shape on the second slide in a PowerPoint presentation:


```vb
With ActivePresentation.Slides(1).Shapes(2) 
 With .Text.Font 
 .Size = 32 
 .Name = "Palatino" 
 .Softedgeformat = msosoftedge6 
 End With 
End With 

```


## See also


#### Concepts


[Object Model Reference](../../Office-Shared-VBA/articles/reference-object-library-reference-for-office.md)

