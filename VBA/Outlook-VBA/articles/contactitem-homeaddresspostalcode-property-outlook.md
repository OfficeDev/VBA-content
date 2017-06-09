---
title: ContactItem.HomeAddressPostalCode Property (Outlook)
keywords: vbaol11.chm1015
f1_keywords:
- vbaol11.chm1015
ms.prod: outlook
api_name:
- Outlook.ContactItem.HomeAddressPostalCode
ms.assetid: 28d65f71-6be6-5d9e-0935-7f09a5f9fa94
ms.date: 06/08/2017
---


# ContactItem.HomeAddressPostalCode Property (Outlook)

Returns or sets a  **String** representing the postal code portion of the home address for the contact. Read/write.


## Syntax

 _expression_ . **HomeAddressPostalCode**

 _expression_ A variable that represents a **ContactItem** object.


## Remarks

This property is parsed from the  **[HomeAddress](contactitem-homeaddress-property-outlook.md)** property, but may be changed or entered independently should it be parsed incorrectly. Note that any such changes or entries to this property will be overwritten by any subsequent changes or entries to **HomeAddress** .


## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

