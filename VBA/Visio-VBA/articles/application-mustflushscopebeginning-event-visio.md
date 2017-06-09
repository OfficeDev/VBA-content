---
title: Application.MustFlushScopeBeginning Event (Visio)
ms.prod: visio
api_name:
- Visio.Application.MustFlushScopeBeginning
ms.assetid: 98a47603-19c0-4588-3d65-1f9d3fe118c1
ms.date: 06/08/2017
---


# Application.MustFlushScopeBeginning Event (Visio)

Occurs before the Microsoft Visio instance is forced to flush its event queue.


## Syntax

Private Sub  _expression_ _**MustFlushScopeBeginning**( **_ByVal app As [IVAPPLICATION]_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _app_|Required| **[IVAPPLICATION]**|The instance of Visio that is forced to flush its event queue.|

## Remarks

This event, along with the  **MustFlushScopeEnded** event, can be used to identify whether an event is being fired because Visio is forced to flush its event queue.

Visio maintains a queue of pending events that it attempts to fire at discrete moments when it is able to process arbitrary requests (callbacks) from event handlers.

Occasionally, Visio is forced to flush its event queue when it is not prepared to handle arbitrary requests. When this occurs, Visio first fires a  **MustFlushScopeBeginning** event, and then it fires the events that are presently in its event queue. After firing all pending events, Visio fires the **MustFlushScopeEnded** event.

After Visio has fired the  **MustFlushScopeBeginning** event, client programs should not call Visio methods that have side effects until the **MustFlushScopeEnded** event is received. A client can perform arbitrary queries of Visio objects when Visio is between the **MustFlushScopeBeginning** event and **MustFlushScopeEnded** event, but operations that cause side effects may fail.

Visio performs a forced flush of its event queue immediately prior to firing a "before" event such as  **BeforeDocumentClose** or **BeforeShapeDelete** because queued events may apply to objects that are about to close or be deleted. Using the **BeforeDocumentClose** event as an example, there can be queued events that apply to a shape object in the document that is being closed. So, before the document closes, Visio fires all the events in its event queue.

When a shape is deleted, events are fired in the following sequence:




1.  **MustFlushScopeBeginning** eventClient should not call methods that have side effects.
    
2. There are zero (0) or more events in the event queue.
    
3.  **BeforeShapeDelete** eventShape is viable, but Visio is going to delete it.
    
4.  **MustFlushScopeEnded** eventClient can resume invoking methods that have side effects.
    
5.  **ShapesDeleted** eventShape has been deleted.
    
6.  **NoEventsPending** eventNo events remain to be fired.
    


An event is fired both before ( **BeforeShapeDeleted** event) and after ( **ShapesDeleted** event) the shape is deleted. If a program monitoring these events requires that additional shapes be deleted in response to the initial shape deletion, it should do so in the **ShapesDeleted** event handler, not the **BeforeShapeDeleted** event handler. The **BeforeShapeDeleted** event is inside the scope of the **MustFlushScopeBeginning** event and the **MustFlushScopeEnded** event, while the **ShapesDeleted** event is not.

 The sequence number of a **MustFlushScopeBeginning** event may be higher than the sequence number of events the client sees after it has received the **MustFlushScopeBeginning** event because Visio assigns sequence numbers to events as they occur. Any events that were queued when the forced flush began have a lower sequence number than the **MustFlushScopeBeginning** event, even though the **MustFlushScopeBeginning** event fires first.

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


