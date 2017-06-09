---
title: SmartArtNode.Nodes Property (Office)
ms.prod: office
api_name:
- Office.SmartArtNode.Nodes
ms.assetid: ed1dc125-5160-ed59-3187-620e3253af59
ms.date: 06/08/2017
---


# SmartArtNode.Nodes Property (Office)

Retrieves the children nodes associated with this Smart Art Node. Read-only


## Syntax

 _expression_. **Nodes**

 _expression_ An expression that returns a **SmartArtNode** object.


## Example

The following code returns the number of nodes in the Smart Art diagram.


```
ActivePresentation.Slides(1).Shapes(1).SmartArtNodes.Nodes.Count
```


## See also


#### Concepts


[SmartArtNode Object](smartartnode-object-office.md)
#### Other resources


[SmartArtNode Object Members](smartartnode-members-office.md)

