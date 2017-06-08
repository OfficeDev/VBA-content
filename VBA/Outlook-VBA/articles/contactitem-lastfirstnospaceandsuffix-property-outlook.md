---
title: ContactItem.LastFirstNoSpaceAndSuffix Property (Outlook)
keywords: vbaol11.chm1082
f1_keywords:
- vbaol11.chm1082
ms.prod: outlook
api_name:
- Outlook.ContactItem.LastFirstNoSpaceAndSuffix
ms.assetid: 15c9527b-3837-d4a0-0249-2cd751e4379f
ms.date: 06/08/2017
---


# ContactItem.LastFirstNoSpaceAndSuffix Property (Outlook)

Returns a  **String** that contains the last name, first name, and suffix of the user without a space. Read-only


## Syntax

 _expression_ . **LastFirstNoSpaceAndSuffix**

 _expression_ A variable that represents a **ContactItem** object.


## Remarks

This property is used only when the  **[FirstName](contactitem-firstname-property-outlook.md)** , **[LastName](contactitem-lastname-property-outlook.md)** , and **[Suffix](contactitem-suffix-property-outlook.md)** properties (the fields that define this property) contain Asian (DBCS) characters. Note that any such changes or entries to the **FirstName** , **LastName** , or **Suffix** properties will be overwritten by any subsequent changes or entries to FullName.


## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

