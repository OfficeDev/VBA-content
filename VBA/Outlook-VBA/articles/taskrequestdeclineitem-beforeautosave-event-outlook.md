---
title: TaskRequestDeclineItem.BeforeAutoSave Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskRequestDeclineItem.BeforeAutoSave
ms.assetid: a1def448-d1cf-3eae-17c9-aeaafa8fd47b
ms.date: 06/08/2017
---


# TaskRequestDeclineItem.BeforeAutoSave Event (Outlook)

Occurs before the item is automatically saved by Outlook.


## Syntax

 _expression_ . **BeforeAutoSave**( **_Cancel_** )

 _expression_ A variable that represents a **TaskRequestDeclineItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **[TaskRequestDeclineItem](taskrequestdeclineitem-object-outlook.md)** to be saved.|

## See also


#### Concepts


[TaskRequestDeclineItem Object](taskrequestdeclineitem-object-outlook.md)

