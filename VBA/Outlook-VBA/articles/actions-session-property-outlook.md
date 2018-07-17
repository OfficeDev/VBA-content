---
title: Actions.Session Property (Outlook)
keywords: vbaol11.chm147
f1_keywords:
- vbaol11.chm147
ms.prod: outlook
api_name:
- Outlook.Actions.Session
ms.assetid: 21792c3f-9669-2f68-7a47-bac172d16620
ms.date: 06/08/2017
---


# Actions.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents an **Actions** object.


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


[Actions Object](actions-object-outlook.md)

