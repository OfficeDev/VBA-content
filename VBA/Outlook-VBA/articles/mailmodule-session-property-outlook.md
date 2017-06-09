---
title: MailModule.Session Property (Outlook)
keywords: vbaol11.chm2814
f1_keywords:
- vbaol11.chm2814
ms.prod: outlook
api_name:
- Outlook.MailModule.Session
ms.assetid: 6b4405e4-c7b8-9837-a494-889e2e17d7ef
ms.date: 06/08/2017
---


# MailModule.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **MailModule** object.


## Remarks

The  **Session** property and the **[GetNamespace](application-getnamespace-method-outlook.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


#### Concepts


[MailModule Object](mailmodule-object-outlook.md)

