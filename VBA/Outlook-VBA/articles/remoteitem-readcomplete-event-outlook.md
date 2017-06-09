---
title: RemoteItem.ReadComplete Event (Outlook)
ms.assetid: 208867c1-b6dc-4ce8-e25a-13a8f6c686ca
ms.date: 06/08/2017
ms.prod: outlook
---


# RemoteItem.ReadComplete Event (Outlook)
Occurs when Outlook has completed reading the properties of the item.

## Version information

Version Added: Outlook 2013 


## Syntax

 _expression_ . **ReadComplete**_(Cancel)_

 _expression_ A variable that represents a **RemoteItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
|||||
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True**, the read operation is not completed and the item is not displayed in the Reading Pane or inspector.|

## Remarks

The  **ReadComplete** event occurs after the[BeforeRead](remoteitem-read-event-outlook.md) event and before the[Read](remoteitem-beforeread-event-outlook.md) event for the item.

To determine when the item is unloaded from memory, use the [Unload](remoteitem-unload-event-outlook.md) event.

The  **ReadComplete** event corresponds to the Exchange Client Extensions (ECE) event **IExchExtMessageEvents::OnReadComplete**.


## See also


#### Concepts


[RemoteItem Object](remoteitem-object-outlook.md)

