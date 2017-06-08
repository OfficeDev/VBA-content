---
title: ContactItem.HomeAddressCountry Property (Outlook)
keywords: vbaol11.chm1014
f1_keywords:
- vbaol11.chm1014
ms.prod: outlook
api_name:
- Outlook.ContactItem.HomeAddressCountry
ms.assetid: a3e1f178-c01c-e7df-ee4e-fc82f89915f0
ms.date: 06/08/2017
---


# ContactItem.HomeAddressCountry Property (Outlook)

Returns or sets a  **String** representing the country/region portion of the home address for the contact. Read/write.


## Syntax

 _expression_ . **HomeAddressCountry**

 _expression_ A variable that represents a **ContactItem** object.


## Remarks

This property is parsed from the  **[HomeAddress](contactitem-homeaddress-property-outlook.md)** property, but may be changed or entered independently should it be parsed incorrectly. Note that any such changes or entries to this property will be overwritten by any subsequent changes or entries to **HomeAddress** .


## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

