---
title: Event.Persistent Property (Visio)
keywords: vis_sdr.chm12614075
f1_keywords:
- vis_sdr.chm12614075
ms.prod: visio
api_name:
- Visio.Event.Persistent
ms.assetid: e8912935-8c85-77ff-4dbc-4394e894af19
ms.date: 06/08/2017
---


# Event.Persistent Property (Visio)

Determines whether an event persists with its document. Read/write.


## Syntax

 _expression_ . **Persistent**

 _expression_ A variable that represents a **Event** object.


### Return Value

Integer


## Remarks

An event is persistable if its action code is  **visActCodeRunAddon** and the event's source object is capable of containing persistent events.

When an event is first created, its  **Persistent** property is set to the same value as its **Persistable** property; if an event can persist, Microsoft Visio assumes it should persist. You can change the initial setting for a persistable event by setting its **Persistent** property to False. However, you cannot change the **Persistent** property of a nonpersistable event—attempting to do so causes an exception.

A nonpersistent event exists as long as a reference is held on the  **Event** object, the **EventList** object that contains the **Event** object, or the source object that has the **EventList** object. When the last reference to any of these objects is released, the nonpersistent event ceases to exist.

A persistent event exists until its  **Event** object is deleted from the source object's **EventList** collection.


 **Note**   Events handled in a Microsoft Visual Basic for Applications (VBA) project are persistent.


