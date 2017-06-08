---
title: InvisibleApp.VisioIsIdle Event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.VisioIsIdle
ms.assetid: 7757a920-6d48-e2ed-db07-dc80be7af566
ms.date: 06/08/2017
---


# InvisibleApp.VisioIsIdle Event (Visio)

Occurs after the application empties its message queue.


## Syntax

Private Sub  _expression_ _**VisioIsIdle**( **_ByVal app As [IVAPPLICATION]_** )

 _expression_ A variable that represents an **InvisibleApp** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _app_|Required| **[IVAPPLICATION]**|The instance of Microsoft Visio that emptied its message queue.|

## Remarks

Visio continually processes messages in its message queue. When its message queue is empty:




1. Visio performs its own idle-time processing.
    
2. Visio tells Microsoft Visual Basic for Applications to perform its idle-time processing.
    
3. If the message queue is still empty, Visio fires the  **VisioIsIdle** event.
    
4. If the message queue is still empty, Visio calls  **WaitMessage** , which is a call to Microsoft Windows that doesn't return until a new message gets added to the Visio message queue.
    


A client program can use the  **VisioIsIdle** event as a signal to perform its own background processing.

The  **VisioIsIdle** event is not the equivalent of a standard timer event. Client programs that need to be called on a periodic basis should use standard timer techniques, because the duration in which Visio is idle (calls **WaitMessage** ) is unpredictable. For client programs that are only monitoring Visio activity, however, the **VisioIsIdle** event can be sufficient, because until **WaitMessage** returns to Visio, there cannot have been any Visio activity since the **VisioIsIdle** event was last fired.

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


