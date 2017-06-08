---
title: DocumentItem.BeforeAutoSave Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.DocumentItem.BeforeAutoSave
ms.assetid: 3aaf57a3-bcc2-d0ba-6fd9-d801452dc4ca
ms.date: 06/08/2017
---


# DocumentItem.BeforeAutoSave Event (Outlook)

Occurs before the item is automatically saved by Outlook.


## Syntax

 _expression_ . **BeforeAutoSave**( **_Cancel_** )

 _expression_ A variable that represents a **DocumentItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **[DocumentItem](documentitem-object-outlook.md)** to be saved.|

## See also


#### Concepts


[DocumentItem Object](documentitem-object-outlook.md)

