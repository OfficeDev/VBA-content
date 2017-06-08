---
title: Inspectors.Session Property (Outlook)
keywords: vbaol11.chm135
f1_keywords:
- vbaol11.chm135
ms.prod: outlook
api_name:
- Outlook.Inspectors.Session
ms.assetid: 32d60741-21f1-39f8-0069-7dddf3078879
ms.date: 06/08/2017
---


# Inspectors.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents an **Inspectors** object.


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


[Inspectors Object](inspectors-object-outlook.md)

