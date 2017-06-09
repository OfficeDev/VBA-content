---
title: ContactItem.OtherAddressCountry Property (Outlook)
keywords: vbaol11.chm1051
f1_keywords:
- vbaol11.chm1051
ms.prod: outlook
api_name:
- Outlook.ContactItem.OtherAddressCountry
ms.assetid: c9fd6c5f-db32-e1d6-1f2f-88c0c12285c7
ms.date: 06/08/2017
---


# ContactItem.OtherAddressCountry Property (Outlook)

Returns or sets a  **String** representing the country/region portion of the other address for the contact. Read/write.


## Syntax

 _expression_ . **OtherAddressCountry**

 _expression_ A variable that represents a **ContactItem** object.


## Remarks

This property is parsed from the  **[OtherAddress](contactitem-otheraddress-property-outlook.md)** property, but may be changed or entered independently should it be parsed incorrectly. Note that any such changes or entries to this property will be overwritten by any subsequent changes or entries to **OtherAddress** .


## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

