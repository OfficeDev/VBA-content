---
title: StorageItem.Subject Property (Outlook)
keywords: vbaol11.chm2151
f1_keywords:
- vbaol11.chm2151
ms.prod: outlook
api_name:
- Outlook.StorageItem.Subject
ms.assetid: 50533838-ad7a-ce4a-4b9e-7923d2868c41
ms.date: 06/08/2017
---


# StorageItem.Subject Property (Outlook)

Returns or sets a  **String** indicating the subject for the Outlook item. Read/write.


## Syntax

 _expression_ . **Subject**

 _expression_ A variable that represents a **StorageItem** object.


## Remarks

This property corresponds to the MAPI property,  **PidTagSubject** . The **Subject** property is the default property for Outlook items.

The  **Subject** serves as a unique identifier for **[StorageItem](storageitem-object-outlook.md)** objects. You should set the subject in a way to ensure that the objects are unique and would not be overwritten by other solution writers. The recommended practice is to use a **ProgID** plus other unique text to identify the **StorageItem**.


## See also


#### Concepts


[StorageItem Object](storageitem-object-outlook.md)

