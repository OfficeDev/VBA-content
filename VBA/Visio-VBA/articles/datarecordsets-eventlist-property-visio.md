---
title: DataRecordsets.EventList Property (Visio)
keywords: vis_sdr.chm16360610
f1_keywords:
- vis_sdr.chm16360610
ms.prod: visio
api_name:
- Visio.DataRecordsets.EventList
ms.assetid: e88ac4c5-f924-7a56-b4e2-dca9772b06d7
ms.date: 06/08/2017
---


# DataRecordsets.EventList Property (Visio)

Returns the  **[EventList](eventlist-object-visio.md)** collection of the **DataRecordsets** collection. Read-only.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **EventList**

 _expression_ An expression that returns a **DataRecordsets** object.


### Return Value

EventList


## Remarks

Once you retrieve the  **EventList** collection, to receive a notification when one of the events in that collection fires, you can pass the ID of the **[Event](event-object-visio.md)** object that represents that event to the **[EventList.AddAdvise ](eventlist-addadvise-method-visio.md)** method as its EventCode parameter.


