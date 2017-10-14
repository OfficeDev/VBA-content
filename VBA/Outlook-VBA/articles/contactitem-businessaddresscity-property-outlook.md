---
title: ContactItem.BusinessAddressCity Property (Outlook)
keywords: vbaol11.chm971
f1_keywords:
- vbaol11.chm971
ms.prod: outlook
api_name:
- Outlook.ContactItem.BusinessAddressCity
ms.assetid: 6c21e0f0-ab9b-5190-6749-4e8f6fc909e8
ms.date: 06/08/2017
---


# ContactItem.BusinessAddressCity Property (Outlook)

Returns or sets a  **String** representing the city name portion of the business address for the contact. Read/write.


## Syntax

 _expression_ . **BusinessAddressCity**

 _expression_ A variable that represents a **ContactItem** object.


## Remarks

This property is parsed from the  **[BusinessAddress](contactitem-businessaddress-property-outlook.md)** property, but may be changed or entered independently should it be parsed incorrectly. Note that any such changes or entries to this property will be overwritten by any subsequent changes or entries to the **BusinessAddress** property.


## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

