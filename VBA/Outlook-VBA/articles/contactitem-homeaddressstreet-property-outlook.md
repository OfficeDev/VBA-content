---
title: ContactItem.HomeAddressStreet Property (Outlook)
keywords: vbaol11.chm1018
f1_keywords:
- vbaol11.chm1018
ms.prod: outlook
api_name:
- Outlook.ContactItem.HomeAddressStreet
ms.assetid: 9a7af500-e817-6fb1-89b4-6b0ef70741bf
ms.date: 06/08/2017
---


# ContactItem.HomeAddressStreet Property (Outlook)

Returns or sets a  **String** representing the street portion of the home address for the contact. Read/write.


## Syntax

 _expression_ . **HomeAddressStreet**

 _expression_ A variable that represents a **ContactItem** object.


## Remarks

This property is parsed from the  **[HomeAddress](contactitem-homeaddress-property-outlook.md)** property, but may be changed or entered independently should it be parsed incorrectly. Note that any such changes or entries to this property will be overwritten by any subsequent changes or entries to **HomeAddress** .


## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

