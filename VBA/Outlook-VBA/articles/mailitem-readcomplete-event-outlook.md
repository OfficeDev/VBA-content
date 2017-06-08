---
title: MailItem.ReadComplete Event (Outlook)
ms.assetid: 39bba654-0683-95a4-9092-3c0ecbbf9104
ms.date: 06/08/2017
ms.prod: outlook
---


# MailItem.ReadComplete Event (Outlook)
Occurs when Outlook has completed reading the properties of the item.

## Version information

Version Added: Outlook 2013 


## Syntax

 _expression_ . **ReadComplete**_(Cancel)_

 _expression_ A variable that represents a **MailItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
|||||
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True**, the read operation is not completed and the item is not displayed in the Reading Pane or inspector.|

## Remarks

The  **ReadComplete** event occurs after the[BeforeRead](mailitem-beforeread-event-outlook.md) event and before the[Read](mailitem-read-event-outlook.md) event for the item.

To determine when the item is unloaded from memory, use the [Unload](mailitem-unload-event-outlook.md) event.

The  **ReadComplete** event corresponds to the Exchange Client Extensions (ECE) event **IExchExtMessageEvents::OnReadComplete**.


## See also


#### Concepts


[MailItem Object](mailitem-object-outlook.md)

