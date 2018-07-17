---
title: TaskRequestItem.BeforeAutoSave Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskRequestItem.BeforeAutoSave
ms.assetid: 0907ec19-5b94-619e-dcd1-8c458294194f
ms.date: 06/08/2017
---


# TaskRequestItem.BeforeAutoSave Event (Outlook)

Occurs before the item is automatically saved by Outlook.


## Syntax

 _expression_ . **BeforeAutoSave**( **_Cancel_** )

 _expression_ A variable that represents a **TaskRequestItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **[TaskRequestItem](taskrequestitem-object-outlook.md)** to be saved.|

## See also


#### Concepts


[TaskRequestItem Object](taskrequestitem-object-outlook.md)

