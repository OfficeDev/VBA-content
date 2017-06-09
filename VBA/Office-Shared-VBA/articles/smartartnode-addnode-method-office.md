---
title: SmartArtNode.AddNode Method (Office)
ms.prod: office
api_name:
- Office.SmartArtNode.AddNode
ms.assetid: f3022423-4416-ab89-ff89-e6c46d65f42c
ms.date: 06/08/2017
---


# SmartArtNode.AddNode Method (Office)

Adds a new SmartArtNode to the data model in the way specified by the SmartArtNodePosition value, and of type SmartArtNodeType.


## Syntax

 _expression_. **AddNode**( **_Position_**, **_Type_** )

 _expression_ An expression that returns a **SmartArtNode** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Position_|Optional|**MsoSmartArtNodePosition**|Specifies the location of the SmartArtNode in the data model. For example,  **msoSmartArtNodeAbove** or **msoSmartArtNodeAfter**.|
| _Type_|Optional|**MsoSmartArtNodeType**|Specifies the type of the added SmartArtNode. For example,  **msoSmartArtNodeTypeAssistant** or **msoSmartArtNodeTypeDefault**.|

### Return Value

SmartArtNode


## Example

The following code adds a default SmartArtNode below the current node. 


```
Dim saNode As SmartArtNode 
 
saNode = saNode.AddNode(msoSmartArtNodeBelow, msoSmartArtNodeTypeDefault)
```


## See also


#### Concepts


[SmartArtNode Object](smartartnode-object-office.md)
#### Other resources


[SmartArtNode Object Members](smartartnode-members-office.md)

