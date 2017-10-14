---
title: FormNameRuleCondition.Session Property (Outlook)
keywords: vbaol11.chm2450
f1_keywords:
- vbaol11.chm2450
ms.prod: outlook
api_name:
- Outlook.FormNameRuleCondition.Session
ms.assetid: ec224a2e-1d45-82d8-0336-9f1f36549b7a
ms.date: 06/08/2017
---


# FormNameRuleCondition.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **FormNameRuleCondition** object.


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


[FormNameRuleCondition Object](formnamerulecondition-object-outlook.md)

