---
title: InvisibleApp.EventInfo Property (Visio)
keywords: vis_sdr.chm17513475
f1_keywords:
- vis_sdr.chm17513475
ms.prod: visio
api_name:
- Visio.InvisibleApp.EventInfo
ms.assetid: a2908ac3-6e92-5e07-5119-97e1d88416ae
ms.date: 06/08/2017
---


# InvisibleApp.EventInfo Property (Visio)

Gets additional information associated with an event, if any exists. Read-only.


## Syntax

 _expression_ . **EventInfo**( **_eventSeqNum_** )

 _expression_ A variable that represents an **InvisibleApp** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _eventSeqNum_|Required| **Long**| **visEvtIDMostRecent** (0) for information about the most recently fired event, or the sequence number of the event to examine.|

### Return Value

String


## Remarks

When Microsoft Visio fires an event, there are a small number of events for which additional information is available. These events are  **BeforeDocumentSaveAs** , **DocumentSavedAs** , **EnterScope** , **ExitScope** , **MarkerEvent** , **ShapesDeleted** , and **ShapeChanged** . Use the application's **EventInfo** property to obtain this information, when available.

The  **EventInfo** property returns the following:




- A string whose contents are specific to the event in question, if the event does record extra information.
    
- An empty string if an event does not record extra information.
    
- An error if Microsoft Visio no longer has information for the specified event.
    


For details about the contents of the  **EventInfo** property for an event, see the specific event topic.

If an event target queries the  **EventInfo** property immediately after being triggered, the most recent event and the event whose sequence number was passed to the target are the same. However, if the target is an add-on implemented by an executable (.exe) file, this may not be the case, because the executable file and Visio are separate tasks that aren't modal with respect to each other.




 **Note**  Event handlers that use the Microsoft Visual Basic for Applications (VBA)  **WithEvents** keyword have access to only the most recent event and must use **visEvtIDMostRecent** .

To ensure that the information returned by the  **EventInfo** property is associated with the same event that triggered the add-on, the executable file can pass <sequence number> as an argument to the **EventInfo** property. You can obtain the sequence number of an event in the following ways:




- If the  **Action** property of the **Event** object returns **visActCodeRunAddon** , the command line string passed to the add-on contains a substring of the form "/eventid=<sequence number>".
    
     **Note**   Even though the substring is labeled "/eventid," don't confuse the <sequence number> passed in the command line string with the **ID** property of the firing **Event** object, which identifies the **Event** object in its **EventList** collection. The number being passed is actually the firing sequence number.
- If the  **Action** property of the **Event** object returns **visActCodeAdvise** , the sequence number is passed as an argument to the **VisEventProc** procedure implemented by the target object.
    



