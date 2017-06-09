---
title: Sequence.AddTriggerEffect Method (PowerPoint)
keywords: vbapp10.chm651013
f1_keywords:
- vbapp10.chm651013
ms.prod: powerpoint
api_name:
- PowerPoint.Sequence.AddTriggerEffect
ms.assetid: 65acf575-5b64-e95c-827d-dada8e915666
ms.date: 06/08/2017
---


# Sequence.AddTriggerEffect Method (PowerPoint)

Adds a trigger effect to the animation in a  **Sequence** object.


## Syntax

 _expression_. **AddTriggerEffect**( **_pShape_**, **_effectId_**, **_trigger_**, **_pTriggerShape_**, **_bookmark_**, **_Level_** )

 _expression_ A variable that represents a **Sequence** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _pShape_|Required|**Shape**|The  **Shape** object with animation.|
| _effectId_|Required|**MsoAnimEffect**|The type of animation.|
| _trigger_|Required|**MsoAnimTriggerType**|The type of trigger effect to add.|
| _pTriggerShape_|Required|**Shape**|The  **Shape** object that represents the trigger.|
| _bookmark_|Optional|**String**|The bookmark.|
| _Level_|Optional|**MsoAnimateByLevel**|The level of animation.|

### Return Value

Effect


## See also


#### Concepts


[Sequence Object](sequence-object-powerpoint.md)

