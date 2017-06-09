---
title: Event.Persistable Property (Visio)
keywords: vis_sdr.chm12614070
f1_keywords:
- vis_sdr.chm12614070
ms.prod: visio
api_name:
- Visio.Event.Persistable
ms.assetid: 3203ac60-ed7f-81cf-6ecf-0095dbc15c48
ms.date: 06/08/2017
---


# Event.Persistable Property (Visio)

Determines whether an event can potentially persist within its document. Read-only.


## Syntax

 _expression_ . **Persistable**

 _expression_ A variable that represents a **Event** object.


### Return Value

Integer


## Remarks

The  **Persistable** property of an **Event** object indicates whether the event can persist, that is, whether the **Event** object can be stored with a Microsoft Visio document between executions of a program. An **Event** object can persist if the following conditions are true:




1. The action code of the  **Event** object must be **visActCodeRunAddon** . If the action code is **visActCodeAdvise** , the event won't persist and must be re-created by a program at run time.
    
2. The source object must be capable of containing persistent events in its  **EventList** collection. The source object's **PersistsEvents** property indicates whether it can contain persistent events. The only source objects currently capable of containing persistent events are **Document** , **Master** , and **Page** objects.
    


If these conditions are met, any of the following events are persistable:




-  **BeforeMasterDelete**
    
-  **BeforePageDelete**
    
-  **BeforeShapeDelete**
    
-  **DocumentOpened**
    
-  **DocumentCreated**
    
-  **MasterAdded**
    
-  **PageAdded**
    


Although an  **Event** object's **Persistable** property indicates whether an event can persist, its **Persistent** property indicates whether that event actually persists. When an **Event** object is first created, its **Persistent** property is set to the same value as its **Persistable** property. That is, a persistable event's **Persistent** property is set to **True** , and a nonpersistable event's **Persistent** property is set to **False** .

A nonpersistent event exists as long as a reference is held on the  **Event** object, the **EventList** object that contains the **Event** object, or the source object that has the **EventList** object. When the last reference to any of these objects is released, the nonpersistent event ceases to exist.

You can change the initial setting for a persistable event by setting its  **Persistent** property to **False** . In this case, the event doesn't persist with its document, even though it could. However, you cannot change the **Persistent** property of a nonpersistent event; attempting to do so will cause an exception.


 **Note**  Events handled in a Microsoft Visual Basic for Applications (VBA) project are persistent.


