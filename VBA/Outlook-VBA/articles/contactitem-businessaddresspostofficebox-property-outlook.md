---
title: ContactItem.BusinessAddressPostOfficeBox Property (Outlook)
keywords: vbaol11.chm974
f1_keywords:
- vbaol11.chm974
ms.prod: outlook
api_name:
- Outlook.ContactItem.BusinessAddressPostOfficeBox
ms.assetid: 447b3e5d-7f8f-372f-d5a6-843ba65a72b7
ms.date: 06/08/2017
---


# ContactItem.BusinessAddressPostOfficeBox Property (Outlook)

Returns or sets a  **String** representing the post office box number portion of the business address for the contact. Read/write.


## Syntax

 _expression_ . **BusinessAddressPostOfficeBox**

 _expression_ A variable that represents a **ContactItem** object.


## Remarks

This property is parsed from the  **[BusinessAddress](contactitem-businessaddress-property-outlook.md)** property, but may be changed or entered independently should it be parsed incorrectly. Note that any such changes or entries to this property will be overwritten by any subsequent changes or entries to the **BusinessAddress** property.


## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

