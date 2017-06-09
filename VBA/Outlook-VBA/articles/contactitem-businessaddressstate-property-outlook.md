---
title: ContactItem.BusinessAddressState Property (Outlook)
keywords: vbaol11.chm975
f1_keywords:
- vbaol11.chm975
ms.prod: outlook
api_name:
- Outlook.ContactItem.BusinessAddressState
ms.assetid: 0d8d9136-6d41-b0ed-f320-6e26fca15cf7
ms.date: 06/08/2017
---


# ContactItem.BusinessAddressState Property (Outlook)

Returns or sets a  **String** representing the state code portion of the business address for the contact. Read/write.


## Syntax

 _expression_ . **BusinessAddressState**

 _expression_ A variable that represents a **ContactItem** object.


## Remarks

This property is parsed from the  **[BusinessAddress](contactitem-businessaddress-property-outlook.md)** property, but may be changed or entered independently should it be parsed incorrectly. Note that any such changes or entries to this property will be overwritten by any subsequent changes or entries to the **BusinessAddress** property.


## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

