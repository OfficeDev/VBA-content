---
title: Event Object (Visio)
keywords: vis_sdr.chm10090
f1_keywords:
- vis_sdr.chm10090
ms.prod: visio
api_name:
- Visio.Event
ms.assetid: f11fffff-2218-8cc4-f223-31d956d1252d
ms.date: 06/08/2017
---


# Event Object (Visio)

A member of the  **EventList** collection of a source object such as a **Document** . An event encapsulates an event code.


## Remarks

An  **Event** object can trigger two kinds of actions: it can run an add-on, or it can send a notification of the event to the calling program. To create an **Event** object, use the **Add** or **AddAdvise** method of an **EventList** object.

The default property of an  **Event** object is **Event** .

The  **Event** property of the **Event** object establishes the event that triggers the action, and its **Action** property indicates the action to be performed.

Use the  **Persistable** property to find out if the event can be stored with a Microsoft Visio document, or the **Persistent** property to find out if the event is stored. Use the **Trigger** method to trigger an **Event** object's action without waiting for the event to occur. Use the **Enabled** property to temporarily disable an event.


