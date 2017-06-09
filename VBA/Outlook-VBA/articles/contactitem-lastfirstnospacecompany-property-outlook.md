---
title: ContactItem.LastFirstNoSpaceCompany Property (Outlook)
keywords: vbaol11.chm1029
f1_keywords:
- vbaol11.chm1029
ms.prod: outlook
api_name:
- Outlook.ContactItem.LastFirstNoSpaceCompany
ms.assetid: 52e60375-954d-ff0d-d06e-9b0fe8823184
ms.date: 06/08/2017
---


# ContactItem.LastFirstNoSpaceCompany Property (Outlook)

Returns a  **String** representing the concatenated last name, first name, and middle name of the contact with no space between the last name and the first name. Read-only.


## Syntax

 _expression_ . **LastFirstNoSpaceCompany**

 _expression_ A variable that represents a **ContactItem** object.


## Remarks

The company name for the contact is included after the middle name. This property is parsed from the  **[CompanyName](contactitem-companyname-property-outlook.md)** , **[LastName](contactitem-lastname-property-outlook.md)** , **[FirstName](contactitem-firstname-property-outlook.md)** , and **[MiddleName](contactitem-middlename-property-outlook.md)** properties. The **LastName** , **FirstName** , and **MiddleName** properties are themselves parsed from the **[FullName](contactitem-fullname-property-outlook.md)** property. The value of this property is only filled when its associated property ( **FirstName** , **LastName** , **MiddleName** , **CompanyName** , and **Suffix** ) contain Asian (DBCS) characters. If the corresponding field does not contain Asian characters, the property will be empty.


## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

