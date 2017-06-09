---
title: ContactItem.MailingAddressPostOfficeBox Property (Outlook)
keywords: vbaol11.chm1038
f1_keywords:
- vbaol11.chm1038
ms.prod: outlook
api_name:
- Outlook.ContactItem.MailingAddressPostOfficeBox
ms.assetid: b4dc4baa-2af8-f008-6f26-3070dd739a6c
ms.date: 06/08/2017
---


# ContactItem.MailingAddressPostOfficeBox Property (Outlook)

Returns or sets a  **String** representing the post office box number portion of the selected mailing address of the contact. Read/write.


## Syntax

 _expression_ . **MailingAddressPostOfficeBox**

 _expression_ A variable that represents a **ContactItem** object.


## Remarks

This property replicates the property indicated by the  **[SelectedMailingAddress](contactitem-selectedmailingaddress-property-outlook.md)** property, which is one of the following **OlMailingAddress** constants: **olBusiness** , **olHome** , **olNone** , or **olOther** . While it can be changed or entered independently, any such changes or entries to this property will be overwritten by any subsequent changes or entries to the property indicated by **SelectedMailingAddress** .


## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

