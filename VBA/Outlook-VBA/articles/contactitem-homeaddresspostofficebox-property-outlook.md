---
title: ContactItem.HomeAddressPostOfficeBox Property (Outlook)
keywords: vbaol11.chm1016
f1_keywords:
- vbaol11.chm1016
ms.prod: outlook
api_name:
- Outlook.ContactItem.HomeAddressPostOfficeBox
ms.assetid: 9c1b310d-13d8-407c-a97e-a52405e37fb2
ms.date: 06/08/2017
---


# ContactItem.HomeAddressPostOfficeBox Property (Outlook)

Returns or sets a  **String** the post office box number portion of the home address for the contact. Read/write.


## Syntax

 _expression_ . **HomeAddressPostOfficeBox**

 _expression_ A variable that represents a **ContactItem** object.


## Remarks

This property is parsed from the  **[HomeAddress](contactitem-homeaddress-property-outlook.md)** property, but may be changed or entered independently should it be parsed incorrectly. Note that any such changes or entries to this property will be overwritten by any subsequent changes or entries to **HomeAddress** .


## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

