---
title: CardView.Session Property (Outlook)
keywords: vbaol11.chm2581
f1_keywords:
- vbaol11.chm2581
ms.prod: outlook
api_name:
- Outlook.CardView.Session
ms.assetid: 2a5b5f88-dc27-ce37-e93b-30c6310fc03b
ms.date: 06/08/2017
---


# CardView.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **CardView** object.


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


[CardView Object](cardview-object-outlook.md)

