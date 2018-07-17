---
title: DistListItem.BeforeDelete Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.DistListItem.BeforeDelete
ms.assetid: 095be9b1-b7c5-bc22-a478-4f4dfeb91641
ms.date: 06/08/2017
---


# DistListItem.BeforeDelete Event (Outlook)

Occurs before an item (which is an instance of the parent object) is deleted.


## Syntax

 _expression_ . **BeforeDelete**( **_Item_** , **_Cancel_** )

 _expression_ A variable that represents a **DistListItem** object.


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


[DistListItem Object](distlistitem-object-outlook.md)

