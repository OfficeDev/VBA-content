---
title: Event.Trigger Method (Visio)
keywords: vis_sdr.chm12651190
f1_keywords:
- vis_sdr.chm12651190
ms.prod: visio
api_name:
- Visio.Event.Trigger
ms.assetid: 093f8ce7-4d8a-c4d6-802f-4dab98fe199e
ms.date: 06/08/2017
---


# Event.Trigger Method (Visio)

Causes an event's action to be performed.


## Syntax

 _expression_ . **Trigger**( **_ContextString_** )

 _expression_ A variable that represents an **Event** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ContextString_|Required| **String**| The string to send to the target of the event.|

### Return Value

Nothing


## Remarks

Triggering an event causes the action associated with the event to be performed. The specified context string is passed to the target of the action:




- If the action is to run an add-on ( **visEvtCodeRunAddon** ), the string is passed in the command line string sent to the add-on.
    
- If the action is to send a notification to the calling program ( **visEvtCodeAdvise** ), the string is passed in the _moreInfo_ parameter of the notification.
    



