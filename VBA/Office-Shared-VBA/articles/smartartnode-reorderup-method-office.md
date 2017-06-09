---
title: SmartArtNode.ReorderUp Method (Office)
ms.prod: office
api_name:
- Office.SmartArtNode.ReorderUp
ms.assetid: 8c33b3cc-3d28-8959-c2ec-6e38ae07fcd2
ms.date: 06/08/2017
---


# SmartArtNode.ReorderUp Method (Office)

Swaps a node with the previous node in the bulleted list. This method reorder's the nodes entire family.


## Syntax

 _expression_. **ReorderUp**

 _expression_ An expression that returns a **SmartArtNode** object.


### Return Value

Nothing


## Remarks

This method simulates clicking the Reorder Up button on the Microsoft Office Fluent Ribbon user interface which is located on the SmartArt Tools tab, on the Design group on Reorder Up.


## Example

The following code swaps the first node with the next node and reorders the nodes parents.


```
ActivePresentation.Slides(1).Shapes(1).SmartArt.Nodes(2).ReorderUp
```


## See also


#### Concepts


[SmartArtNode Object](smartartnode-object-office.md)
#### Other resources


[SmartArtNode Object Members](smartartnode-members-office.md)

