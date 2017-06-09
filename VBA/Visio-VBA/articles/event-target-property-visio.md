---
title: Event.Target Property (Visio)
keywords: vis_sdr.chm12614485
f1_keywords:
- vis_sdr.chm12614485
ms.prod: visio
api_name:
- Visio.Event.Target
ms.assetid: 92e78a1d-5888-9984-a3c6-6e39ac15c18b
ms.date: 06/08/2017
---


# Event.Target Property (Visio)

Gets or sets the target of an event. Read/write.


## Syntax

 _expression_ . **Target**

 _expression_ A variable that represents a **Event** object.


### Return Value

String


## Remarks

An event consists of an event-action pair. When the event occurs, the action is performed. An event also specifies the target of the action and arguments to send to the target.

If the action code of the event is  **visActCodeRunAddon** , the **Target** property contains the name of the add-on to run.

If the action code of the event is  **visActCodeAdvise** , the **Target** property is not available. Attempting to get or set the **Target** property for such an event causes an exception.


