---
title: EventList Object (Visio)
keywords: vis_sdr.chm10095
f1_keywords:
- vis_sdr.chm10095
ms.prod: visio
api_name:
- Visio.EventList
ms.assetid: 08b70863-ce73-2cd2-ccc0-a993bd261ea2
ms.date: 06/08/2017
---


# EventList Object (Visio)

Includes an  **Event** object for each event to which an object should respond. The object that possesses the event list is sometimes called the source object.


## Remarks

To retrieve an  **EventList** collection, use the **EventList** property of the source object.

The default property of  **EventList** is **Item** .

In general, the level of the source object in the Microsoft Visio object hierarchy determines the scope of its response. For example, if an  **Event** object for the **DocumentOpened** event is in the **EventList** collection of a **Document** object, that event's action is triggered only when that document is opened. If the same **Event** object is in the **EventList** collection of an **Application** object, the event's action is triggered whenever any document is opened in that instance of Visio.

To create an  **Event** object that runs an add-on, use the **Add** method of an **EventList** collection.

To create an  **Event** object that sends a notification, use the **AddAdvise** method.


