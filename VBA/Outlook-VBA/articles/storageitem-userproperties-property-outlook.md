---
title: StorageItem.UserProperties Property (Outlook)
keywords: vbaol11.chm2149
f1_keywords:
- vbaol11.chm2149
ms.prod: outlook
api_name:
- Outlook.StorageItem.UserProperties
ms.assetid: 0a08e77c-1665-a612-2f47-ef1c3fc331d2
ms.date: 06/08/2017
---


# StorageItem.UserProperties Property (Outlook)

Returns the  **[UserProperties](userproperties-object-outlook.md)** collection that represents all the user properties for the Outlook item. Read-only.


## Syntax

 _expression_ . **UserProperties**

 _expression_ A variable that represents a **StorageItem** object.


## Remarks

If you use the  **[UserProperties.Add](userproperties-add-method-outlook.md)** method on the **[UserProperties](userproperties-object-outlook.md)** object associated with a **[StorageItem](storageitem-object-outlook.md)** , the optional _AddToFolderFields_ and _DisplayFormat_ arguments of the **UserProperties.Add** method will be ignored. Any custom properties of the **StorageItem** object will not be exposed as custom properties in the **Field Chooser**.


## See also


#### Concepts


[StorageItem Object](storageitem-object-outlook.md)

