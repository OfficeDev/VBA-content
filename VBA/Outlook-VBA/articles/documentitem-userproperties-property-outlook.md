---
title: DocumentItem.UserProperties Property (Outlook)
keywords: vbaol11.chm1208
f1_keywords:
- vbaol11.chm1208
ms.prod: outlook
api_name:
- Outlook.DocumentItem.UserProperties
ms.assetid: c2253136-5b4d-4f27-e7b5-93ed96b0f76f
ms.date: 06/08/2017
---


# DocumentItem.UserProperties Property (Outlook)

Returns the  **[UserProperties](userproperties-object-outlook.md)** collection that represents all the user properties for the Outlook item. Read-only.


## Syntax

 _expression_ . **UserProperties**

 _expression_ A variable that represents a **DocumentItem** object.


## Remarks

Even though  **olWordDocumentItem** is a valid constant in the **[OlItemType](olitemtype-enumeration-outlook.md)** enumeration, user-defined fields cannot to be added to a **[DocumentItem](documentitem-object-outlook.md)** object and you will receive an error when you try to programmatically add a user-defined field to a **DocumentItem** object.


## See also


#### Concepts


[DocumentItem Object](documentitem-object-outlook.md)

