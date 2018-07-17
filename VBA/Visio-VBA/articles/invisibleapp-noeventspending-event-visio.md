---
title: InvisibleApp.NoEventsPending Event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.NoEventsPending
ms.assetid: 65947eae-69de-3220-e4e5-5edf5b6ad242
ms.date: 06/08/2017
---


# InvisibleApp.NoEventsPending Event (Visio)

Occurs after the Microsoft Visio instance flushes its event queue.


## Syntax

Private Sub  _expression_ _**NoEventsPending**( **_ByVal app As [IVAPPLICATION]_** )

 _expression_ A variable that represents an **InvisibleApp** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _app_|Required| **[IVAPPLICATION]**|The instance of Visio that flushed its event queue.|

## Remarks

Visio maintains a queue of events and fires them at discrete moments. Immediately after Visio fires the last event in its event queue, it fires a  **NoEventsPending** event.

A client program can use the  **NoEventsPending** event as a signal that Visio has completed a burst of activity. For example, a client program may want to react to changes in a shape's geometry. A single user action performed on the shape can generate several **CellChanged** events. The client program could record selected information for each **CellChanged** event and perform its processing after it receives the **NoEventsPending** event.

Visio fires the  **NoEventsPending** event only if at least one of the events in the queue is being listened to. If no program is listening for any of the queued events, the **NoEventsPending** event does not fire. If your program is only listening to the **NoEventsPending** event, it does not fire unless another program is listening for some of the queued events.

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


