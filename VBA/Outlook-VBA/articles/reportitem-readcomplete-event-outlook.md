---
title: ReportItem.ReadComplete Event (Outlook)
ms.assetid: f73cb164-0c88-f439-6474-a4502b6731ea
ms.date: 06/08/2017
ms.prod: outlook
---


# ReportItem.ReadComplete Event (Outlook)
Occurs when Outlook has completed reading the properties of the item.

## Version information

Version Added: Outlook 2013 


## Syntax

 _expression_ . **ReadComplete**_(Cancel)_

 _expression_ A variable that represents a **ReportItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
|||||
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True**, the read operation is not completed and the item is not displayed in the Reading Pane or inspector.|

## Remarks

The  **ReadComplete** event occurs after the[BeforeRead](reportitem-beforeread-event-outlook.md) event and before the[Read](reportitem-read-event-outlook.md) event for the item.

To determine when the item is unloaded from memory, use the [Unload](reportitem-unload-event-outlook.md) event.

The  **ReadComplete** event corresponds to the Exchange Client Extensions (ECE) event **IExchExtMessageEvents::OnReadComplete**.


## See also


#### Concepts


[ReportItem Object](reportitem-object-outlook.md)

