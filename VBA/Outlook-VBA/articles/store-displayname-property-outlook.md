---
title: Store.DisplayName Property (Outlook)
keywords: vbaol11.chm800
f1_keywords:
- vbaol11.chm800
ms.prod: outlook
api_name:
- Outlook.Store.DisplayName
ms.assetid: 785ec583-3553-6002-41b6-d0c6d0028b5a
ms.date: 06/08/2017
---


# Store.DisplayName Property (Outlook)

Returns a  **String** representing the display name of the **[Store](store-object-outlook.md)** object. Read-only.


## Syntax

 _expression_ . **DisplayName**

 _expression_ A variable that represents a **Store** object.


## Remarks

 **DisplayName** is the default property of the **Store** object. This property corresponds to the MAPI property, **PidTagDisplayName** .

 **DisplayName** is read-only. To change the **DisplayName** of a Personal Folders File (.pst), use the **[PropertyAccessor](propertyaccessor-object-outlook.md)** object and the **[PropertyAccessor.SetProperty](propertyaccessor-setproperty-method-outlook.md)** method.


## See also


#### Concepts


[Store Object](store-object-outlook.md)

