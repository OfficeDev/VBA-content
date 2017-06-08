---
title: ContactItem.BeforeAutoSave Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.ContactItem.BeforeAutoSave
ms.assetid: c9fe9c4d-3c00-455c-3e89-9ac584597117
ms.date: 06/08/2017
---


# ContactItem.BeforeAutoSave Event (Outlook)

Occurs before the item is automatically saved by Outlook.


## Syntax

 _expression_ . **BeforeAutoSave**( **_Cancel_** , )

 _expression_ A variable that represents a **ContactItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **[ContactItem](contactitem-object-outlook.md)** to be saved.|

## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

