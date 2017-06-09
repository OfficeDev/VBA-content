---
title: MailItem.BeforeAutoSave Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.MailItem.BeforeAutoSave
ms.assetid: 0c725b91-f72f-7ceb-b2a9-da4f0369cf41
ms.date: 06/08/2017
---


# MailItem.BeforeAutoSave Event (Outlook)

Occurs before the item is automatically saved by Outlook.


## Syntax

 _expression_ . **BeforeAutoSave**( **_Cancel_** , )

 _expression_ A variable that represents a **MailItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **[MailItem](mailitem-object-outlook.md)** to be saved.|

## See also


#### Concepts


[MailItem Object](mailitem-object-outlook.md)

