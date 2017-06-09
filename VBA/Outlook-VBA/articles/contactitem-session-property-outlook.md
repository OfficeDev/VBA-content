---
title: ContactItem.Session Property (Outlook)
keywords: vbaol11.chm928
f1_keywords:
- vbaol11.chm928
ms.prod: outlook
api_name:
- Outlook.ContactItem.Session
ms.assetid: b67eb0d4-9b97-2be7-fc24-ecdd58fb01ca
ms.date: 06/08/2017
---


# ContactItem.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **ContactItem** object.


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


[ContactItem Object](contactitem-object-outlook.md)

