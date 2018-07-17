---
title: ExchangeDistributionList.Session Property (Outlook)
keywords: vbaol11.chm2110
f1_keywords:
- vbaol11.chm2110
ms.prod: outlook
api_name:
- Outlook.ExchangeDistributionList.Session
ms.assetid: 9488e161-d297-d999-538d-a8b295380701
ms.date: 06/08/2017
---


# ExchangeDistributionList.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents an **ExchangeDistributionList** object.


## Remarks

The  **Session** property and the **[Application.GetNamespace](application-getnamespace-method-outlook.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


#### Concepts


[ExchangeDistributionList Object](exchangedistributionlist-object-outlook.md)

