---
title: AnimationBehaviors.Add Method (PowerPoint)
keywords: vbapp10.chm656004
f1_keywords:
- vbapp10.chm656004
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationBehaviors.Add
ms.assetid: 427e7faa-1fc7-a145-98bc-1954054c2aec
ms.date: 06/08/2017
---


# AnimationBehaviors.Add Method (PowerPoint)

Returns an  **[AnimationBehavior](animationbehavior-object-powerpoint.md)** object that represents a new animation behavior.


## Syntax

 _expression_. **Add**( **_Type_**, **_Index_** )

 _expression_ A variable that represents an **AnimationBehaviors** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Required|**MsoAnimType**|The type of the animation behavior.|
| _Index_|Optional|**Long**|The position of the animation behaviorce in relation to other animation behaviors. The default value is -1, which means that if you omit the  _Index_ parameter, the new animation behavior is added at the end of the existing animation behaviors.|

### Return Value

AnimationBehavior


## Remarks

The  _Type_ parameter value can be one of these **MsoAnimType** constants.


||
|:-----|
|**msoAnimTypeColor**|
|**msoAnimTypeMixed**|
|**msoAnimTypeMotion**|
|**msoAnimTypeNone**|
|**msoAnimTypeProperty**|
|**msoAnimTypeRotation**|
|**msoAnimTypeScale**|

## See also


#### Concepts


[AnimationBehaviors Object](animationbehaviors-object-powerpoint.md)

