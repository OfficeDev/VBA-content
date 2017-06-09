---
title: FromRssFeedRuleCondition.Session Property (Outlook)
keywords: vbaol11.chm3255
f1_keywords:
- vbaol11.chm3255
ms.prod: outlook
api_name:
- Outlook.FromRssFeedRuleCondition.Session
ms.assetid: 72939751-3012-fdc9-dfb7-60306bc522cd
ms.date: 06/08/2017
---


# FromRssFeedRuleCondition.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **FromRssFeedRuleCondition** object.


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


[FromRssFeedRuleCondition Object](fromrssfeedrulecondition-object-outlook.md)

