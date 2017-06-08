---
title: Event.Action Property (Visio)
keywords: vis_sdr.chm12613010
f1_keywords:
- vis_sdr.chm12613010
ms.prod: visio
api_name:
- Visio.Event.Action
ms.assetid: dd776f54-051c-13c3-433e-299687203381
ms.date: 06/08/2017
---


# Event.Action Property (Visio)

Gets or sets the action code of an  **Event** object. Read/write.


## Syntax

 _expression_ . **Action**

 _expression_ A variable that represents a **Event** object.


### Return Value

Integer


## Remarks

An  **Event** object consists of an event-action pair; an event triggers an action. An action code is the numeric constant for the action that the event triggers.

Microsoft Visio supports the following action codes.



|**Constant **|**Value **|
|:-----|:-----|
| **visActCodeRunAddon**|1 |
| **visActCodeAdvise**|2 |

