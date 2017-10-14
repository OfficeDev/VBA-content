---
title: ContactItem.ReadComplete Event (Outlook)
ms.assetid: 1700ad85-3113-e937-9eb3-be78246fd4d5
ms.date: 06/08/2017
ms.prod: outlook
---


# ContactItem.ReadComplete Event (Outlook)
Occurs when Outlook has completed reading the properties of the item.

## Version information

Version Added: Outlook 2013 


## Syntax

 _expression_ . **ReadComplete**_(Cancel)_

 _expression_ A variable that represents a **ContactItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
|||||
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True**, the read operation is not completed and the item is not displayed in the Reading Pane or inspector.|

## Remarks

The  **ReadComplete** event occurs after the[BeforeRead](contactitem-beforeread-event-outlook.md) event and before the[Read](contactitem-read-event-outlook.md) event for the item.

To determine when the item is unloaded from memory, use the [Unload](contactitem-unload-event-outlook.md) event.

The  **ReadComplete** event corresponds to the Exchange Client Extensions (ECE) event **IExchExtMessageEvents::OnReadComplete**.


## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

