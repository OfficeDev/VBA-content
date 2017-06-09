---
title: ContactItem.HomeAddressCity Property (Outlook)
keywords: vbaol11.chm1013
f1_keywords:
- vbaol11.chm1013
ms.prod: outlook
api_name:
- Outlook.ContactItem.HomeAddressCity
ms.assetid: 1d2334f2-0401-3bcc-53bf-fa55e1664d9c
ms.date: 06/08/2017
---


# ContactItem.HomeAddressCity Property (Outlook)

Returns or sets a  **String** representing the city portion of the home address for the contact. Read/write.


## Syntax

 _expression_ . **HomeAddressCity**

 _expression_ A variable that represents a **ContactItem** object.


## Remarks

This property is parsed from the  **[HomeAddress](contactitem-homeaddress-property-outlook.md)** property, but may be changed or entered independently should it be parsed incorrectly. Note that any such changes or entries to this property will be overwritten by any subsequent changes or entries to **HomeAddress** .


## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

