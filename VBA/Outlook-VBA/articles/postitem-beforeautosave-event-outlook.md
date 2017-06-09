---
title: PostItem.BeforeAutoSave Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.PostItem.BeforeAutoSave
ms.assetid: 61a44326-0215-869b-0824-2308fd8017cf
ms.date: 06/08/2017
---


# PostItem.BeforeAutoSave Event (Outlook)

Occurs before the item is automatically saved by Outlook.


## Syntax

 _expression_ . **BeforeAutoSave**( **_Cancel_** )

 _expression_ A variable that represents a **PostItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **[PostItem](postitem-object-outlook.md)** to be saved.|

## See also


#### Concepts


[PostItem Object](postitem-object-outlook.md)

