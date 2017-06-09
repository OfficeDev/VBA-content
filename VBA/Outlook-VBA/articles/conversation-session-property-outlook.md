---
title: Conversation.Session Property (Outlook)
keywords: vbaol11.chm3386
f1_keywords:
- vbaol11.chm3386
ms.prod: outlook
api_name:
- Outlook.Conversation.Session
ms.assetid: 6f41faaa-e16a-d171-ed72-d2fef64a77f1
ms.date: 06/08/2017
---


# Conversation.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **[Conversation](conversation-object-outlook.md)** object.


## Remarks

This property returns  **Null** ( **Nothing** in Visual Basic) if there is no logged-on session.

You can use the  **Session** property and the **[GetNamespace](application-getnamespace-method-outlook.md)** method interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:




```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```




```vb
Set objSession = Application.Session
```


## See also


#### Concepts


[Conversation Object](conversation-object-outlook.md)

