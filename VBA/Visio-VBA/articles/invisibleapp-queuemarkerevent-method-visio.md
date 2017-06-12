---
title: InvisibleApp.QueueMarkerEvent Method (Visio)
keywords: vis_sdr.chm17516455
f1_keywords:
- vis_sdr.chm17516455
ms.prod: visio
api_name:
- Visio.InvisibleApp.QueueMarkerEvent
ms.assetid: ed782045-49b1-dcab-de81-41a45117afe7
ms.date: 06/08/2017
---


# InvisibleApp.QueueMarkerEvent Method (Visio)

Queues a  **MarkerEvent** event that fires after all other queued events.


## Syntax

 _expression_ . **QueueMarkerEvent**( **_ContextString_** , **_lpi4Ret_** )

 _expression_ A variable that represents an **InvisibleApp** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ContextString_|Required| **String**|An arbitrary string that is passed with the event that fires.|

### Return Value

Long


## Remarks

The  **QueueMarkerEvent** method works in conjunction with the **MarkerEvent** event to allow an Automation client to queue an event to itself. The **QueueMarkerEvent** method causes the application to fire a **MarkerEvent** event after it has fired all the events in its event queue.

The  **QueueMarkerEvent** method returns the sequence number of the **MarkerEvent** event to fire, and the string passed to the **QueueMarkerEvent** method (legally empty) is passed to the **MarkerEvent** event handler.

A client program can use either the sequence number or the string to correlate  **QueueMarkerEvent** calls with **MarkerEvent** events. In this way, the client is able to distinguish events it caused and events it did not cause.


## Example

Paste this example code into the  **ThisDocument** object and then run the **UseMarker** procedure. The output will be displayed in the Microsoft Visual Basic for Applications (VBA) Immediate window.


```vb
 
Dim WithEvents vsoApplication As Visio.Application 
 
Private Sub vsoApplication_MarkerEvent(ByVal app As Visio.IVApplication, _ 
 ByVal SequenceNum As Long, ByVal ContextString As String) 
 
 Debug.Print "Marker: " &; app.EventInfo(0) 
 
End Sub 
 
Private Sub vsoApplication_ShapeAdded(ByVal Shape As Visio.IVShape) 
 
 Debug.Print " ShapeAdded: " &; Shape.Name 
 
End Sub 
 
Public Sub UseMarker() 
 
 Set vsoApplication = ThisDocument.Application 
 
 'Marker events can be used to comment a segment 
 'of events in the queue. 
 vsoApplication.QueueMarkerEvent "I am starting..." 
 ActivePage.DrawRectangle 0, 0, 3, 3 
 vsoApplication.QueueMarkerEvent "I am finished..." 
 
End Sub
```

The output in the VBA Immediate window looks like this:

Marker: I am starting...

ShapeAdded: Sheet.1

Marker: I am finished...


