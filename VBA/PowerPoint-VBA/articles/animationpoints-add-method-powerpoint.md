---
title: AnimationPoints.Add Method (PowerPoint)
keywords: vbapp10.chm663004
f1_keywords:
- vbapp10.chm663004
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationPoints.Add
ms.assetid: faa75675-aac2-af60-b3a3-a142181ef10b
ms.date: 06/08/2017
---


# AnimationPoints.Add Method (PowerPoint)

Returns an  **[AnimationPoint](animationpoint-object-powerpoint.md)** object that represents a new animation point.


## Syntax

 _expression_. **Add**( **_Index_** )

 _expression_ A variable that represents an **AnimationPoints** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Optional|**Long**|The position of the animation point in relation to other animation points. The default value is -1, which means that if you omit the Index parameter, the new animation point is added to the end of existing animation points.|

### Return Value

AnimationPoint


## See also


#### Concepts


[AnimationPoints Object](animationpoints-object-powerpoint.md)

