---
title: TaskItem.ReadComplete Event (Outlook)
ms.assetid: 0706a4b9-1035-bdf9-a48d-8d039a2001fa
ms.date: 06/08/2017
ms.prod: outlook
---


# TaskItem.ReadComplete Event (Outlook)
Occurs when Outlook has completed reading the properties of the item.

## Version information

Version Added: Outlook 2013 


## Syntax

 _expression_ . **ReadComplete**_(Cancel)_

 _expression_ A variable that represents a **TaskItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
|||||
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True**, the read operation is not completed and the item is not displayed in the Reading Pane or inspector.|

## Remarks

The  **ReadComplete** event occurs after the[BeforeRead](taskitem-beforeread-event-outlook.md) event and before the[Read](taskitem-read-event-outlook.md) event for the item.

To determine when the item is unloaded from memory, use the [Unload](taskitem-unload-event-outlook.md) event.

The  **ReadComplete** event corresponds to the Exchange Client Extensions (ECE) event **IExchExtMessageEvents::OnReadComplete**.


## See also


#### Concepts


[TaskItem Object](taskitem-object-outlook.md)

