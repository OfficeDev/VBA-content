---
title: RuleActions.Count Property (Outlook)
keywords: vbaol11.chm2182
f1_keywords:
- vbaol11.chm2182
ms.prod: outlook
api_name:
- Outlook.RuleActions.Count
ms.assetid: 91b4425f-0e17-fff1-0d9c-1697b205ff2a
ms.date: 06/08/2017
---


# RuleActions.Count Property (Outlook)

Returns a  **Long** indicating the count of objects in the specified collection. Read-only.


## Syntax

 _expression_ . **Count**

 _expression_ A variable that represents a **RuleActions** object.


## Remarks

You can enumerate the  **[RuleActions](ruleactions-object-outlook.md)** collection from 1 through **RuleActions.Count** to determine all the rule actions for an existing rule. Although the Rules OM supports creating rules with only the most commonly used rule actions and not all rule actions supported by the Rules and Alerts Wizard, the **RuleActions** collection includes all rule actions of a rule. Hence you can always enumerate the **RuleActions** collection object to determine which rule actions are enabled for an existing rule.


## See also


#### Concepts


[RuleActions Object](ruleactions-object-outlook.md)

