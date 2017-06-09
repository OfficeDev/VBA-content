---
title: ContactItem.MailingAddressPostalCode Property (Outlook)
keywords: vbaol11.chm1037
f1_keywords:
- vbaol11.chm1037
ms.prod: outlook
api_name:
- Outlook.ContactItem.MailingAddressPostalCode
ms.assetid: bdb1cd44-1ae5-598d-0f25-604deafdb7ed
ms.date: 06/08/2017
---


# ContactItem.MailingAddressPostalCode Property (Outlook)

Returns or sets a  **String** representing the postal code (zip code) portion of the selected mailing address of the contact. Read/write.


## Syntax

 _expression_ . **MailingAddressPostalCode**

 _expression_ A variable that represents a **ContactItem** object.


## Remarks

This property replicates the property indicated by the  **[SelectedMailingAddress](contactitem-selectedmailingaddress-property-outlook.md)** property, which is one of the following **OlMailingAddress** constants: **olBusiness** , **olHome** , **olNone** , or **olOther** . While it can be changed or entered independently, any such changes or entries to this property will be overwritten by any subsequent changes or entries to the property indicated by **SelectedMailingAddress** .


## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

