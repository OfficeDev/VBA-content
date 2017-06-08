---
title: StorageItem.Creator Property (Outlook)
keywords: vbaol11.chm2152
f1_keywords:
- vbaol11.chm2152
ms.prod: outlook
api_name:
- Outlook.StorageItem.Creator
ms.assetid: c89c777c-5f4b-f672-ff74-d34db3bcd790
ms.date: 06/08/2017
---


# StorageItem.Creator Property (Outlook)

Returns and sets the solution that created the  **[StorageItem](storageitem-object-outlook.md)** object. Read/write.


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **StorageItem** object.


## Remarks

Outlook does not set the  **Creator** property. Use the **Creator** property to identify the **StorageItem** objects you have created for your add-in. One recomended value for this property is the programmatic identifier (ProgID) of the add-in.


## See also


#### Concepts


[StorageItem Object](storageitem-object-outlook.md)

