---
title: TransitionPeriod Property
keywords: fm20.chm5225109
f1_keywords:
- fm20.chm5225109
ms.prod: office
api_name:
- Office.TransitionPeriod
ms.assetid: cfdda5c3-244b-4404-d6a8-544755056473
ms.date: 06/08/2017
---


# TransitionPeriod Property



Specifies the duration, in milliseconds, of a transition effect.
 **Syntax**
 _object_. **TransitionPeriod** [= _Long_ ]
The  **TransitionPeriod** property syntax has these parts:


| <strong>Part</strong> | <strong>Description</strong>                                                     |
|:----------------------|:---------------------------------------------------------------------------------|
| <em>object</em>       | Required. A valid object.                                                        |
| <em>Long</em>         | Optional. How long it takes to complete the transition from one page to another. |

 **Remarks**
Any integer from zero to 10000 is a valid setting for this property. Setting the  **TransitionPeriod** property to zero disables the transition effect; setting **TransitionPeriod** to 10000 creates a 10-second transition.

