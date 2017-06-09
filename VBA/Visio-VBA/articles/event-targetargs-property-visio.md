---
title: Event.TargetArgs Property (Visio)
keywords: vis_sdr.chm12614490
f1_keywords:
- vis_sdr.chm12614490
ms.prod: visio
api_name:
- Visio.Event.TargetArgs
ms.assetid: b2102b52-de0d-30f2-042c-5ebdbf7aaffd
ms.date: 06/08/2017
---


# Event.TargetArgs Property (Visio)

Gets or sets the arguments to be sent to the target of an event. Read/write.


## Syntax

 _expression_ . **TargetArgs**

 _expression_ A variable that represents a **Event** object.


### Return Value

String


## Remarks

An event consists of an event-action pair. When the event occurs, the action is performed. An event also specifies the target of the action and arguments to send to the target.

When you use  **visActCodeRunAddon** , the **TargetArgs** property contains the arguments to send to the add-on when it is run.

When you use  **visActCodeAdvise** , the **TargetArgs** property contains the string specified with the **AddAdvise** method when the **Event** object was created. When the program receives notification of the event, it can get the **Event** object and its **TargetArgs** property to obtain the string.


