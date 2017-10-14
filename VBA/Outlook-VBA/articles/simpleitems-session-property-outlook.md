---
title: SimpleItems.Session Property (Outlook)
keywords: vbaol11.chm3395
f1_keywords:
- vbaol11.chm3395
ms.prod: outlook
api_name:
- Outlook.SimpleItems.Session
ms.assetid: 5445d76f-658c-babf-87cf-44efd75a208a
ms.date: 06/08/2017
---


# SimpleItems.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **[SimpleItems](simpleitems-object-outlook.md)** object.


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


[SimpleItems Object](simpleitems-object-outlook.md)

