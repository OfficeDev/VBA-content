---
title: SharingItem.BeforeDelete Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.SharingItem.BeforeDelete
ms.assetid: 60726a1b-2d74-c7a6-fef8-b26f5f5e7d01
ms.date: 06/08/2017
---


# SharingItem.BeforeDelete Event (Outlook)

Occurs before an item (which is an instance of the parent object) is deleted.


## Syntax

 _expression_ . **BeforeDelete**( **_Item_** , **_Cancel_** )

 _expression_ An expression that returns a **SharingItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Item_|Required| **Object**|The item being deleted.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the operation is not completed and the item is not deleted.|

## Remarks

In order for this event to fire when a sharing message is deleted through an action, an inspector must be open.

The event occurs each time an item is deleted.


## See also


#### Concepts


[SharingItem Object](sharingitem-object-outlook.md)

