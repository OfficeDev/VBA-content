---
title: RemoteItem.BeforeAutoSave Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.RemoteItem.BeforeAutoSave
ms.assetid: f33e1442-0e65-cc78-34ac-496b65ba565e
ms.date: 06/08/2017
---


# RemoteItem.BeforeAutoSave Event (Outlook)

Occurs before the item is automatically saved by Outlook.


## Syntax

 _expression_ . **BeforeAutoSave**( **_Cancel_** )

 _expression_ A variable that represents a **RemoteItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **[RemoteItem](remoteitem-object-outlook.md)** to be saved.|

## See also


#### Concepts


[RemoteItem Object](remoteitem-object-outlook.md)

