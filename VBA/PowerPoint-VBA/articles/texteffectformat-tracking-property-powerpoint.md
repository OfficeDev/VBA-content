---
title: TextEffectFormat.Tracking Property (PowerPoint)
keywords: vbapp10.chm556014
f1_keywords:
- vbapp10.chm556014
ms.prod: powerpoint
api_name:
- PowerPoint.TextEffectFormat.Tracking
ms.assetid: 998cbec0-959c-e76f-9e26-6e8466894a60
ms.date: 06/08/2017
---


# TextEffectFormat.Tracking Property (PowerPoint)

Returns or sets the ratio of the horizontal space allotted to each character in the specified text to the width of the character. Read/write. 


## Syntax

 _expression_. **Tracking**

 _expression_ A variable that represents a **TextEffectFormat** object.


### Return Value

Single


## Remarks

The  **Tracking** property value can be from 0 (zero) through 5. (Large values for this property specify ample space between characters; values less than 1 can produce character overlap.)

The following table gives the values of the  **Tracking** property that correspond to the settings available in the user interface.



|**User interface setting**|**Equivalent Tracking property value**|
|:-----|:-----|
|Very Tight|0.99925|
|Tight|0.999925|
|Normal|1.0|
|Loose|1.003|
|Very Loose|1.006|

## See also


#### Concepts


[TextEffectFormat Object](texteffectformat-object-powerpoint.md)

