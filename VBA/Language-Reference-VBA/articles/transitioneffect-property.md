---
title: TransitionEffect Property
keywords: fm20.chm5225108
f1_keywords:
- fm20.chm5225108
ms.prod: office
api_name:
- Office.TransitionEffect
ms.assetid: 10a65973-fa2e-5b9d-5052-ead41286e1af
ms.date: 06/08/2017
---


# TransitionEffect Property



Specifies the visual effect to use when changing from one page to another.
 **Syntax**
 _object_. **TransitionEffect** [= _fmTransitionEffect_ ]
The  **TransitionEffect** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _fmTransitionEffect_|Optional. The transition effect you want between pages.|
 **Settings**
The settings for  _fmTransitionEffect_ are:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| _fmTransitionEffectNone_|0|No special effect (default).|
| _fmTransitionEffectCoverUp_|1|The new page covers the old page, moving from the bottom to the top.|
| _fmTransitionEffectCoverRightUp_|2|The new page covers the old page, moving from the bottom left corner to the top right corner.|
| _fmTransitionEffectCoverRight_|3|The new page covers the old page, moving from the left edge to the right.|
| _fmTransitionEffectCoverRightDown_|4|The new page covers the old page, moving from the top left corner to the bottom right corner.|
| _fmTransitionEffectCoverDown_|5|The new page covers the old page, moving from the top to the bottom.|
| _fmTransitionEffectCoverLeftDown_|6|The new page covers the old page, moving from the top right corner to the bottom left corner.|
| _fmTransitionEffectCoverLeft_|7|The new page covers the old page, moving from the right to the left.|
| _fmTransitionEffectCoverLeftUp_|8|The new page covers the old page, moving from the bottom right corner to the top left corner.|
| _fmTransitionEffectPushUp_|9|The new page pushes the old page out of view, moving from the bottom to the top.|
| _fmTransitionEffectPushRight_|10|The new page pushes the old page out of view, moving from the left to the right.|
| _fmTransitionEffectPushDown_|11|The new page pushes the old page out of view, moving from the top to the bottom.|
| _fmTransitionEffectPushLeft_|12|The new page pushes the old page out of view, moving from the right to the left.|
 **Remarks**
Use the  **TransitionPeriod** property to specify the duration of a transition effect.

