---
title: Explorer.Session Property (Outlook)
keywords: vbaol11.chm2759
f1_keywords:
- vbaol11.chm2759
ms.prod: outlook
api_name:
- Outlook.Explorer.Session
ms.assetid: 47752d87-6ef5-4838-4c08-0325c0b613f7
ms.date: 06/08/2017
---


# Explorer.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents an **Explorer** object.


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


[Explorer Object](explorer-object-outlook.md)

