---
title: ContactItem.HomeAddressState Property (Outlook)
keywords: vbaol11.chm1017
f1_keywords:
- vbaol11.chm1017
ms.prod: outlook
api_name:
- Outlook.ContactItem.HomeAddressState
ms.assetid: bc052902-1e38-3d6a-1b7b-308861357731
ms.date: 06/08/2017
---


# ContactItem.HomeAddressState Property (Outlook)

Returns or sets a  **String** representing the state portion of the home address for the contact. Read/write.


## Syntax

 _expression_ . **HomeAddressState**

 _expression_ A variable that represents a **ContactItem** object.


## Remarks

This property is parsed from the  **[HomeAddress](contactitem-homeaddress-property-outlook.md)** property, but may be changed or entered independently should it be parsed incorrectly. Note that any such changes or entries to this property will be overwritten by any subsequent changes or entries to **HomeAddress** .


## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

