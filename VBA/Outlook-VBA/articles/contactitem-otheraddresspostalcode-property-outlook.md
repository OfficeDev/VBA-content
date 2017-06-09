---
title: ContactItem.OtherAddressPostalCode Property (Outlook)
keywords: vbaol11.chm1052
f1_keywords:
- vbaol11.chm1052
ms.prod: outlook
api_name:
- Outlook.ContactItem.OtherAddressPostalCode
ms.assetid: a9cecb5e-d6c3-9496-8537-fab14520321f
ms.date: 06/08/2017
---


# ContactItem.OtherAddressPostalCode Property (Outlook)

Returns or sets a  **String** representing the postal code portion of the other address for the contact. Read/write.


## Syntax

 _expression_ . **OtherAddressPostalCode**

 _expression_ A variable that represents a **ContactItem** object.


## Remarks

This property is parsed from the  **[OtherAddress](contactitem-otheraddress-property-outlook.md)** property, but may be changed or entered independently should it be parsed incorrectly. Note that any such changes or entries to this property will be overwritten by any subsequent changes or entries to **OtherAddress** .


## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

