---
title: Rules.Session Property (Outlook)
keywords: vbaol11.chm2156
f1_keywords:
- vbaol11.chm2156
ms.prod: outlook
api_name:
- Outlook.Rules.Session
ms.assetid: c544e009-623c-3e4d-b71a-9177dcfcc668
ms.date: 06/08/2017
---


# Rules.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **Rules** object.


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


[Rules Object](rules-object-outlook.md)

