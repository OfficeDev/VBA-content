---
title: ContactItem.CompanyLastFirstSpaceOnly Property (Outlook)
keywords: vbaol11.chm985
f1_keywords:
- vbaol11.chm985
ms.prod: outlook
api_name:
- Outlook.ContactItem.CompanyLastFirstSpaceOnly
ms.assetid: 8f78b5c8-3832-8c30-6ba6-d7f0149d2dd3
ms.date: 06/08/2017
---


# ContactItem.CompanyLastFirstSpaceOnly Property (Outlook)

Returns a  **String** representing the company name for the contact followed by the concatenated last name, first name, and middle name with spaces between the last, first, and middle names. Read-only.


## Syntax

 _expression_ . **CompanyLastFirstSpaceOnly**

 _expression_ A variable that represents a **ContactItem** object.


## Remarks

This property is parsed from the  **[CompanyName](contactitem-companyname-property-outlook.md)** , **[LastName](contactitem-lastname-property-outlook.md)** , **[FirstName](contactitem-firstname-property-outlook.md)** , and **[MiddleName](contactitem-middlename-property-outlook.md)** properties. The **LastName** , **FirstName** , and **MiddleName** properties are themselves parsed from the **[FullName](contactitem-fullname-property-outlook.md)** property. The value of this property is only filled when its associated property ( **FirstName** , **LastName** , **MiddleName** , **CompanyName** , and **Suffix** ) contain Asian (DBCS) characters. If the corresponding field does not contain Asian characters, the property will be empty.


## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

