---
title: MailItem.Session Property (Outlook)
keywords: vbaol11.chm1292
f1_keywords:
- vbaol11.chm1292
ms.prod: outlook
api_name:
- Outlook.MailItem.Session
ms.assetid: 43272ff5-ab89-f160-7995-981158f6f375
ms.date: 06/08/2017
---


# MailItem.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **MailItem** object.


## Remarks

The  **Session** property and the **[GetNamespace](application-getnamespace-method-outlook.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


#### Concepts


[MailItem Object](mailitem-object-outlook.md)

