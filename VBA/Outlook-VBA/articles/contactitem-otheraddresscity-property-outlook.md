---
title: ContactItem.OtherAddressCity Property (Outlook)
keywords: vbaol11.chm1050
f1_keywords:
- vbaol11.chm1050
ms.prod: outlook
api_name:
- Outlook.ContactItem.OtherAddressCity
ms.assetid: ab29f816-1434-658b-196b-a918a4234aa7
ms.date: 06/08/2017
---


# ContactItem.OtherAddressCity Property (Outlook)

Returns or sets a  **String** representing the city portion of the other address for the contact. Read/write.


## Syntax

 _expression_ . **OtherAddressCity**

 _expression_ A variable that represents a **ContactItem** object.


## Remarks

This property is parsed from the  **[OtherAddress](contactitem-otheraddress-property-outlook.md)** property, but may be changed or entered independently should it be parsed incorrectly. Note that any such changes or entries to this property will be overwritten by any subsequent changes or entries to **OtherAddress** .


## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

