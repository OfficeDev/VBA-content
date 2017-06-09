---
title: SmartArtNode.Level Property (Office)
ms.prod: office
api_name:
- Office.SmartArtNode.Level
ms.assetid: 63143dbc-ecd2-240c-f4c1-2b32cd47872d
ms.date: 06/08/2017
---


# SmartArtNode.Level Property (Office)

Retrieves the node's level in the hierarchy. Read-only


## Syntax

 _expression_. **Level**

 _expression_ An expression that returns a **SmartArtNode** object.


## Remarks

The levels start at 1 and increment upward. If a node has no level, then a 0 is returned. For example, in the following data model, A and F have a level of 1, B and D have a level of 2, and C and E have a level of 3.


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
    

## See also


#### Concepts


[SmartArtNode Object](smartartnode-object-office.md)
#### Other resources


[SmartArtNode Object Members](smartartnode-members-office.md)

