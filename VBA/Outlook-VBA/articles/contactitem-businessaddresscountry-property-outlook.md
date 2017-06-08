---
title: ContactItem.BusinessAddressCountry Property (Outlook)
keywords: vbaol11.chm972
f1_keywords:
- vbaol11.chm972
ms.prod: outlook
api_name:
- Outlook.ContactItem.BusinessAddressCountry
ms.assetid: cd5b1640-ddbd-9fca-062c-f03ed39f7821
ms.date: 06/08/2017
---


# ContactItem.BusinessAddressCountry Property (Outlook)

Returns or sets a  **String** representing the country/region code portion of the business address for the contact. Read/write.


## Syntax

 _expression_ . **BusinessAddressCountry**

 _expression_ A variable that represents a **ContactItem** object.


## Remarks

This property is parsed from the  **[BusinessAddress](contactitem-businessaddress-property-outlook.md)** property, but may be changed or entered independently should it be parsed incorrectly. Note that any such changes or entries to this property will be overwritten by any subsequent changes or entries to the **BusinessAddress** property.


## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

