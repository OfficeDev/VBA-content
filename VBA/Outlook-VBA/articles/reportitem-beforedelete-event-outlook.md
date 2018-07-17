---
title: ReportItem.BeforeDelete Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.ReportItem.BeforeDelete
ms.assetid: 2fca7e89-39b3-73c4-715a-003921a055cd
ms.date: 06/08/2017
---


# ReportItem.BeforeDelete Event (Outlook)

Occurs before an item (which is an instance of the parent object) is deleted.


## Syntax

 _expression_ . **BeforeDelete**( **_Item_** , **_Cancel_** )

 _expression_ A variable that represents a **ReportItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Item_|Required| **Object**|The item being deleted.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the operation is not completed and the item is not deleted.|

## Remarks

In order for this event to fire when an e-mail message, distribution list, journal entry, task, contact, or post are deleted through an action, an inspector must be open.

The event occurs each time an item is deleted.


## See also


#### Concepts


[ReportItem Object](reportitem-object-outlook.md)

