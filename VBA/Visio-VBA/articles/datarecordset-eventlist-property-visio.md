---
title: DataRecordset.EventList Property (Visio)
keywords: vis_sdr.chm16460610
f1_keywords:
- vis_sdr.chm16460610
ms.prod: visio
api_name:
- Visio.DataRecordset.EventList
ms.assetid: 419cdd3d-cb12-cbb6-5e47-d343b1a84d74
ms.date: 06/08/2017
---


# DataRecordset.EventList Property (Visio)

Returns the  **[EventList](eventlist-object-visio.md)** collection of the **DataRecordset** object. Read-only.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **EventList**

 _expression_ An expression that returns a **DataRecordset** object.


### Return Value

EventList


## Remarks

Once you retrieve the  **EventList** collection, to receive a notification when one of the events in that collection fires, you can pass the ID of the **[Event](event-object-visio.md)** object that represents that event to the **[EventList.AddAdvise ](eventlist-addadvise-method-visio.md)** method as its EventCode parameter.


