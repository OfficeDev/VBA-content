---
title: Folders.Session Property (Outlook)
keywords: vbaol11.chm41
f1_keywords:
- vbaol11.chm41
ms.prod: outlook
api_name:
- Outlook.Folders.Session
ms.assetid: 1f8d8e11-d4d9-6769-37af-5c97e1413023
ms.date: 06/08/2017
---


# Folders.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **Folders** object.


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


[Folders Object](folders-object-outlook.md)

