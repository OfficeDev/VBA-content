---
title: SmartArtNodes.Add Method (Office)
ms.prod: office
api_name:
- Office.SmartArtNodes.Add
ms.assetid: 51254d1a-0395-2b40-842c-84ba3d52a98b
ms.date: 06/08/2017
---


# SmartArtNodes.Add Method (Office)

Adds a new  **SmartArtNode** object to the diagram with specified text.


## Syntax

 _expression_. **Add**

 _expression_ An expression that returns a **SmartArtNodes** object.


### Return Value

SmartArtNode


## Remarks

This new node is added to the bottom of the data model at the top most level for this collection of nodes. If the highest level was 2, then the new node's level would then be 2 as well. 


## Example

The following code adds a SmartArtNode to the collection of SmartArtNodes. 


```
Dim saNodes As SmartArtNodes 
saNodes.Add()
```


## See also


#### Concepts


[SmartArtNodes Object](smartartnodes-object-office.md)
#### Other resources


[SmartArtNodes Object Members](smartartnodes-members-office.md)

