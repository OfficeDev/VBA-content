---
title: ContactItem.OtherAddressPostOfficeBox Property (Outlook)
keywords: vbaol11.chm1053
f1_keywords:
- vbaol11.chm1053
ms.prod: outlook
api_name:
- Outlook.ContactItem.OtherAddressPostOfficeBox
ms.assetid: 905500a2-475a-ed2a-79b5-e46a3d8c117c
ms.date: 06/08/2017
---


# ContactItem.OtherAddressPostOfficeBox Property (Outlook)

Returns or sets a  **String** representing the post office box portion of the other address for the contact. Read/write.


## Syntax

 _expression_ . **OtherAddressPostOfficeBox**

 _expression_ A variable that represents a **ContactItem** object.


## Remarks

This property is parsed from the  **[OtherAddress](contactitem-otheraddress-property-outlook.md)** property, but may be changed or entered independently should it be parsed incorrectly. Note that any such changes or entries to this property will be overwritten by any subsequent changes or entries to **OtherAddress** .


## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

