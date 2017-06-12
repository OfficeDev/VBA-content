---
title: SharingItem.BeforeAutoSave Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.SharingItem.BeforeAutoSave
ms.assetid: 38515dda-2539-5f0b-4c04-831067c09327
ms.date: 06/08/2017
---


# SharingItem.BeforeAutoSave Event (Outlook)

Occurs before the  **[SharingItem](sharingitem-object-outlook.md)** is automatically saved by Outlook.


## Syntax

 _expression_ . **BeforeAutoSave**( **_Cancel_** )

 _expression_ An expression that returns a **SharingItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **SharingItem** to be saved.|

## See also


#### Concepts


[SharingItem Object](sharingitem-object-outlook.md)

