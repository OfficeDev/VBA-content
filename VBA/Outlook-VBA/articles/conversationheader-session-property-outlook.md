---
title: ConversationHeader.Session Property (Outlook)
keywords: vbaol11.chm3548
f1_keywords:
- vbaol11.chm3548
ms.prod: outlook
api_name:
- Outlook.ConversationHeader.Session
ms.assetid: 1262a068-ad5f-492d-2a96-edc365956fe6
ms.date: 06/08/2017
---


# ConversationHeader.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **[ConversationHeader](conversationheader-object-outlook.md)** object.


## Remarks

Returns  **Null** ( **Nothing** in Visual Basic) if there is no logged-on session.

You can use the  **Session** property and the **[GetNamespace](application-getnamespace-method-outlook.md)** method interchangeably to obtain the **NameSpace** object for the current session.


## See also


#### Concepts


[ConversationHeader Object](conversationheader-object-outlook.md)

