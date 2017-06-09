---
title: AppointmentItem.ReadComplete Event (Outlook)
ms.assetid: 749e8d58-c15c-0b63-5486-cc2aa2190638
ms.date: 06/08/2017
ms.prod: outlook
---


# AppointmentItem.ReadComplete Event (Outlook)
Occurs when Outlook has completed reading the properties of the item.

## Syntax

 _expression_ . **ReadComplete**_(Cancel)_

 _expression_ A variable that represents an **AppointmentItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
|||||
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True**, the read operation is not completed and the item is not displayed in the Reading Pane or inspector.|

## Remarks

The  **ReadComplete** event occurs after the[BeforeRead](appointmentitem-beforeread-event-outlook.md) event and before the[Read](appointmentitem-read-event-outlook.md) event for the item.

To determine when the item is unloaded from memory, use the [Unload](appointmentitem-unload-event-outlook.md) event.

The  **ReadComplete** event corresponds to the Exchange Client Extensions (ECE) event **IExchExtMessageEvents::OnReadComplete**.


## See also


#### Concepts


[AppointmentItem Object](appointmentitem-object-outlook.md)

