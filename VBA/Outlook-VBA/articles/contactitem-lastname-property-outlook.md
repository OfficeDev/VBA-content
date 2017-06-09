---
title: ContactItem.LastName Property (Outlook)
keywords: vbaol11.chm1032
f1_keywords:
- vbaol11.chm1032
ms.prod: outlook
api_name:
- Outlook.ContactItem.LastName
ms.assetid: 430682f6-a230-887b-404b-a71989121fa2
ms.date: 06/08/2017
---


# ContactItem.LastName Property (Outlook)

Returns or sets a  **String** representing the last name for the contact. Read/write.


## Syntax

 _expression_ . **LastName**

 _expression_ A variable that represents a **ContactItem** object.


## Remarks

This property is parsed from the  **[FullName](contactitem-fullname-property-outlook.md)** property, but may be changed or entered independently should it be parsed incorrectly. Note that any such changes or entries to this property will be overwritten by any subsequent changes of entries to **FullName** .


## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

