---
title: AccountRuleCondition.Session Property (Outlook)
keywords: vbaol11.chm2379
f1_keywords:
- vbaol11.chm2379
ms.prod: outlook
api_name:
- Outlook.AccountRuleCondition.Session
ms.assetid: 1bcc0f04-a3a1-40e5-5853-938e284db89f
ms.date: 06/08/2017
---


# AccountRuleCondition.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents an **AccountRuleCondition** object.


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


[AccountRuleCondition Object](accountrulecondition-object-outlook.md)

