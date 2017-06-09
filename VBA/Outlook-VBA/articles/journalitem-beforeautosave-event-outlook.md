---
title: JournalItem.BeforeAutoSave Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.JournalItem.BeforeAutoSave
ms.assetid: b4924fd8-52cd-fa8d-11d8-2683ea2f5b52
ms.date: 06/08/2017
---


# JournalItem.BeforeAutoSave Event (Outlook)

Occurs before the item is automatically saved by Outlook.


## Syntax

 _expression_ . **BeforeAutoSave**( **_Cancel_** , )

 _expression_ A variable that represents a **JournalItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **[JournalItem](journalitem-object-outlook.md)** to be saved.|

## See also


#### Concepts


[JournalItem Object](journalitem-object-outlook.md)

