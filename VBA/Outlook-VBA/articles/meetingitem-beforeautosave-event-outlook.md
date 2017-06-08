---
title: MeetingItem.BeforeAutoSave Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.MeetingItem.BeforeAutoSave
ms.assetid: 59de272e-a36a-e842-a962-03ebe2befa26
ms.date: 06/08/2017
---


# MeetingItem.BeforeAutoSave Event (Outlook)

Occurs before the item is automatically saved by Outlook.


## Syntax

 _expression_ . **BeforeAutoSave**( **_Cancel_** )

 _expression_ A variable that represents a **MeetingItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **[MeetingItem](meetingitem-object-outlook.md)** to be saved.|

## See also


#### Concepts


[MeetingItem Object](meetingitem-object-outlook.md)

