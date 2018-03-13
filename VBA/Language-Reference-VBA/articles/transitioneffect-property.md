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


| <strong>Part</strong>       | <strong>Description</strong>                            |
|:----------------------------|:--------------------------------------------------------|
| <em>object</em>             | Required. A valid object.                               |
| <em>fmTransitionEffect</em> | Optional. The transition effect you want between pages. |

 **Settings**
The settings for  _fmTransitionEffect_ are:


| <strong>Constant</strong>                 | <strong>Value</strong> | <strong>Description</strong>                                                                  |
|:------------------------------------------|:-----------------------|:----------------------------------------------------------------------------------------------|
| <em>fmTransitionEffectNone</em>           | 0                      | No special effect (default).                                                                  |
| <em>fmTransitionEffectCoverUp</em>        | 1                      | The new page covers the old page, moving from the bottom to the top.                          |
| <em>fmTransitionEffectCoverRightUp</em>   | 2                      | The new page covers the old page, moving from the bottom left corner to the top right corner. |
| <em>fmTransitionEffectCoverRight</em>     | 3                      | The new page covers the old page, moving from the left edge to the right.                     |
| <em>fmTransitionEffectCoverRightDown</em> | 4                      | The new page covers the old page, moving from the top left corner to the bottom right corner. |
| <em>fmTransitionEffectCoverDown</em>      | 5                      | The new page covers the old page, moving from the top to the bottom.                          |
| <em>fmTransitionEffectCoverLeftDown</em>  | 6                      | The new page covers the old page, moving from the top right corner to the bottom left corner. |
| <em>fmTransitionEffectCoverLeft</em>      | 7                      | The new page covers the old page, moving from the right to the left.                          |
| <em>fmTransitionEffectCoverLeftUp</em>    | 8                      | The new page covers the old page, moving from the bottom right corner to the top left corner. |
| <em>fmTransitionEffectPushUp</em>         | 9                      | The new page pushes the old page out of view, moving from the bottom to the top.              |
| <em>fmTransitionEffectPushRight</em>      | 10                     | The new page pushes the old page out of view, moving from the left to the right.              |
| <em>fmTransitionEffectPushDown</em>       | 11                     | The new page pushes the old page out of view, moving from the top to the bottom.              |
| <em>fmTransitionEffectPushLeft</em>       | 12                     | The new page pushes the old page out of view, moving from the right to the left.              |

 **Remarks**
Use the  **TransitionPeriod** property to specify the duration of a transition effect.

