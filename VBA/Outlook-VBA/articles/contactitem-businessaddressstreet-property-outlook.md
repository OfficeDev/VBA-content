---
title: ContactItem.BusinessAddressStreet Property (Outlook)
keywords: vbaol11.chm976
f1_keywords:
- vbaol11.chm976
ms.prod: outlook
api_name:
- Outlook.ContactItem.BusinessAddressStreet
ms.assetid: 1d3e67c4-b02d-c2cf-b04b-85bc1464d788
ms.date: 06/08/2017
---


# ContactItem.BusinessAddressStreet Property (Outlook)

Returns or sets a  **String** representing the street address portion of the business address for the contact. Read/write.


## Syntax

 _expression_ . **BusinessAddressStreet**

 _expression_ A variable that represents a **ContactItem** object.


## Remarks

This property is parsed from the  **[BusinessAddress](contactitem-businessaddress-property-outlook.md)** property, but may be changed or entered independently should it be parsed incorrectly. Note that any such changes or entries to this property will be overwritten by any subsequent changes or entries to the **BusinessAddress** property.


## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

