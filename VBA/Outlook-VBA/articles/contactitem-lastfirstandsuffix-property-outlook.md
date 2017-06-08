---
title: ContactItem.LastFirstAndSuffix Property (Outlook)
keywords: vbaol11.chm1027
f1_keywords:
- vbaol11.chm1027
ms.prod: outlook
api_name:
- Outlook.ContactItem.LastFirstAndSuffix
ms.assetid: b234614c-e2c0-cba2-6ec8-69be1a31caf1
ms.date: 06/08/2017
---


# ContactItem.LastFirstAndSuffix Property (Outlook)

Returns a  **String** representing the last name, first name, middle name, and suffix of the contact. Read-only.


## Syntax

 _expression_ . **LastFirstAndSuffix**

 _expression_ A variable that represents a **ContactItem** object.


## Remarks

There is a comma between the last and first names and spaces between all the names and the suffix. This property is parsed from the  **[LastName](contactitem-lastname-property-outlook.md)** , **[FirstName](contactitem-firstname-property-outlook.md)** , **[MiddleName](contactitem-middlename-property-outlook.md)** and **[Suffix](contactitem-suffix-property-outlook.md)** properties. The **LastName** , **FirstName** , and **Suffix** properties are themselves parsed from the **[FullName](contactitem-fullname-property-outlook.md)** property. The value of this property is only filled when its associated property ( **FirstName** , **LastName** , **MiddleName** , **CompanyName** , and **Suffix** ) contain Asian (DBCS) characters. If the corresponding field does not contain Asian characters, the property will be empty.


## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

