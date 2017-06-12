---
title: ContactItem.FirstName Property (Outlook)
keywords: vbaol11.chm1004
f1_keywords:
- vbaol11.chm1004
ms.prod: outlook
api_name:
- Outlook.ContactItem.FirstName
ms.assetid: 403b5e5a-037b-cf21-efc2-2bd2a80c3789
ms.date: 06/08/2017
---


# ContactItem.FirstName Property (Outlook)

Returns or sets a  **String** representing the first name for the contact. Read/write.


## Syntax

 _expression_ . **FirstName**

 _expression_ A variable that represents a **ContactItem** object.


## Remarks

This property is parsed from the  **[FullName](contactitem-fullname-property-outlook.md)** property, but may be changed or entered independently should it be parsed incorrectly. Note that any such changes or entries to this property will be overwritten by any subsequent changes of entries to **FullName** .


## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

