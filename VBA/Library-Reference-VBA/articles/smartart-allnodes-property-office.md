---
title: SmartArt.AllNodes Property (Office)
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.SmartArt.AllNodes
ms.assetid: 8562a464-61dd-e019-9f44-89ade4703589
---


# SmartArt.AllNodes Property (Office)

Retrieves a  **SmartArtNodes** object containing all of the nodes within the SmartArt diagram. Read-only


## Syntax

 _expression_. **AllNodes**

 _expression_ An expression that returns a **SmartArt** object.


## Remarks

The nodes are retrieved in order, independent of data model. For example, the following data model would retrieve the nodes in order A, B, C, D, E, F.


- A
    
- 
      - B
    
  - 
      - C
    
  - D
    
- 
      - 
      - E
    
- F
    

## Example

The following example sets the text inside the first node.


```
smartart.AllNodes(1).TextFrame2.TextRange.Text="Node 1"
```


## See also


#### Concepts


[SmartArt Object](smartart-object-office.md)

