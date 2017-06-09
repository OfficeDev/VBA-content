---
title: Application.EventsEnabled Property (Visio)
keywords: vis_sdr.chm10013485
f1_keywords:
- vis_sdr.chm10013485
ms.prod: visio
api_name:
- Visio.Application.EventsEnabled
ms.assetid: 92775825-c17d-fd4f-a315-7a181d75aed5
ms.date: 06/08/2017
---


# Application.EventsEnabled Property (Visio)

Determines whether a Microsoft Visio instance fires events. Read/write.


## Syntax

 _expression_ . **EventsEnabled**

 _expression_ A variable that represents an **Application** object.


### Return Value

Integer


## Remarks

If the  **EventsEnabled** property is **False** , Visio does not fire events, run add-ons, or execute strings that contain arbitrary Visual Basic for Applications (VBA) code when evaluating RUNADDON operands in cell formulas.

By default, the  **EventsEnabled** property is **True** when an instance of Visio starts.

You may want to disable event firing if you have written code to handle events such as  **DocumentOpened** or **DocumentCreated** that does not work properly, or to prevent the incorporation of a virus into a document. Events will not fire until the **EventsEnabled** property is set to **True** .


## Example

These VBA macros show how to use the  **EventsEnabled** property to suspend and resume event processing.


```vb
 
Public Sub SuspendEventProcessing_Example() 
 
 'Suspend event processing. 
 Application.EventsEnabled = False 
 End Sub 
 
Public Sub EventsEnabled_Example() 
 
 'Resume event processing. 
 Application.EventsEnabled = True 
 
End Sub
```


