---
title: SmartArtNode.Promote Method (Office)
ms.prod: office
api_name:
- Office.SmartArtNode.Promote
ms.assetid: 806dae89-7a05-7597-70dc-ad297c79fbff
ms.date: 06/08/2017
---


# SmartArtNode.Promote Method (Office)

Promotes the current node (and all its children) a single level within the data model.


## Syntax

 _expression_. **Promote**

 _expression_ An expression that returns a **SmartArtNode** object.


### Return Value

Nothing


## Remarks

This functionality mimics the promote button on the Microsoft Office Fluent Ribbon UI when working within the content pane. For example, given the following data model: if B is promoted, the resulting data model looks like the following: 


- A
    
- 
      - B
    
  - 
      - C
    
- D
    

- A
    
- B
    
- 
      - C
    
- D
    

## See also


#### Concepts


[SmartArtNode Object](smartartnode-object-office.md)
#### Other resources


[SmartArtNode Object Members](smartartnode-members-office.md)

