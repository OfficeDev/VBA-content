---
title: ContactItem.MailingAddressState Property (Outlook)
keywords: vbaol11.chm1039
f1_keywords:
- vbaol11.chm1039
ms.prod: outlook
api_name:
- Outlook.ContactItem.MailingAddressState
ms.assetid: 9e15bba8-2256-fd1a-60ae-ac63d6d4f4e3
ms.date: 06/08/2017
---


# ContactItem.MailingAddressState Property (Outlook)

Returns or sets a  **String** representing the state code portion for the selected mailing address of the contact. Read/write.


## Syntax

 _expression_ . **MailingAddressState**

 _expression_ A variable that represents a **ContactItem** object.


## Remarks

This property replicates the property indicated by the  **[SelectedMailingAddress](contactitem-selectedmailingaddress-property-outlook.md)** property, which is one of the following **OlMailingAddress** constants: **olBusiness** , **olHome** , **olNone** , or **olOther** . While it can be changed or entered independently, any such changes or entries to this property will be overwritten by any subsequent changes or entries to the property indicated by **SelectedMailingAddress** .


## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

