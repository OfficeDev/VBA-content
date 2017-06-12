---
title: OlStorageIdentifierType Enumeration (Outlook)
keywords: vbaol11.chm3111
f1_keywords:
- vbaol11.chm3111
ms.prod: outlook
api_name:
- Outlook.OlStorageIdentifierType
ms.assetid: 14283b38-6a0d-2954-bffe-87c36af27b2c
ms.date: 06/08/2017
---


# OlStorageIdentifierType Enumeration (Outlook)

Specifies the type of identifier for a  **[StorageItem](storageitem-object-outlook.md)** object.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **olIdentifyByEntryID**|1|Identifies a  **StorageItem** by **[EntryID](storageitem-entryid-property-outlook.md)** .|
| **olIdentifyByMessageClass**|2|Identifies a  **StorageItem** by message class.|
| **olIdentifyBySubject**|0|Identifies a  **StorageItem** by **[Subject](storageitem-subject-property-outlook.md)** .|

## Remarks

The message class of a [StorageItem Object (Outlook)](storageitem-object-outlook.md) is not exposed as an explicit built-in property. You can access the message class property through the[PropertyAccessor Object (Outlook)](propertyaccessor-object-outlook.md) that is provided by[StorageItem.PropertyAccessor Property (Outlook)](storageitem-propertyaccessor-property-outlook.md).


