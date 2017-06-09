---
title: Folder.Session Property (Outlook)
keywords: vbaol11.chm1983
f1_keywords:
- vbaol11.chm1983
ms.prod: outlook
api_name:
- Outlook.Folder.Session
ms.assetid: b446d857-4f41-085f-7303-3e5050e33bba
ms.date: 06/08/2017
---


# Folder.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **Folder** object.


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


[Folder Object](folder-object-outlook.md)

