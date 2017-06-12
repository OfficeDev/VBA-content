---
title: ImportanceRuleCondition.Session Property (Outlook)
keywords: vbaol11.chm2334
f1_keywords:
- vbaol11.chm2334
ms.prod: outlook
api_name:
- Outlook.ImportanceRuleCondition.Session
ms.assetid: 521d650f-8724-e8cb-6d20-1e7d730bf419
ms.date: 06/08/2017
---


# ImportanceRuleCondition.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents an **ImportanceRuleCondition** object.


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


[ImportanceRuleCondition Object](importancerulecondition-object-outlook.md)

