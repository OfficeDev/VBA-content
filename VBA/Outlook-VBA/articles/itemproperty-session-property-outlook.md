---
title: ItemProperty.Session Property (Outlook)
keywords: vbaol11.chm520
f1_keywords:
- vbaol11.chm520
ms.prod: outlook
api_name:
- Outlook.ItemProperty.Session
ms.assetid: f33cfcd0-f86b-d0cd-7d35-a21644bc5c42
ms.date: 06/08/2017
---


# ItemProperty.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents an **ItemProperty** object.


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


[ItemProperty Object](itemproperty-object-outlook.md)

