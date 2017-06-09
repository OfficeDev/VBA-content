---
title: JournalItem.ReadComplete Event (Outlook)
ms.assetid: 63f74eb2-99bc-2ce7-c412-c28eba80e75c
ms.date: 06/08/2017
ms.prod: outlook
---


# JournalItem.ReadComplete Event (Outlook)
Occurs when Outlook has completed reading the properties of the item.

## Version information

Version Added: Outlook 2013 


## Syntax

 _expression_ . **ReadComplete**_(Cancel)_

 _expression_ A variable that represents a **JournalItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
|||||
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True**, the read operation is not completed and the item is not displayed in the Reading Pane or inspector.|

## Remarks

The  **ReadComplete** event occurs after the[BeforeRead](journalitem-beforeread-event-outlook.md) event and before the[Read](journalitem-read-event-outlook.md) event for the item.

To determine when the item is unloaded from memory, use the [Unload](journalitem-unload-event-outlook.md) event.

The  **ReadComplete** event corresponds to the Exchange Client Extensions (ECE) event **IExchExtMessageEvents::OnReadComplete**.


## See also


#### Concepts


[JournalItem Object](journalitem-object-outlook.md)

