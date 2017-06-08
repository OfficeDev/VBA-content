---
title: Rule.Actions Property (Outlook)
keywords: vbaol11.chm2174
f1_keywords:
- vbaol11.chm2174
ms.prod: outlook
api_name:
- Outlook.Rule.Actions
ms.assetid: 2b1e2ad4-c735-b3a8-6b27-5004f10393ce
ms.date: 06/08/2017
---


# Rule.Actions Property (Outlook)

Returns a  **[RuleActions](ruleactions-object-outlook.md)** collection object that represents all the available rule actions for the rule. Read-only.


## Syntax

 _expression_ . **Actions**

 _expression_ A variable that represents a **Rule** object.


## Remarks

You can enumerate and enable rules with any rule action that the Rules and Alerts Wizard support, but you can programmatically create rules that have only the most commonly used rule actions, and not any rule action that the Rules and Alerts Wizard supports. For more information on rule action support, see [Specify Rule Actions](http://msdn.microsoft.com/library/c5f83c81-0e01-38aa-5ec7-3932b4443e43%28Office.15%29.aspx).

Through the  **Actions** property, each rule is associated with a **RuleActions** object. The **RuleActions** collection is a fixed object - you cannot add or remove items from this collection. Rule actions that are enabled in the rule will have an enabled rule action in the **RuleActions** collection. Rule actions that are not enabled in the rule will have a rule action in this collection that has the **[RuleAction.Enabled](ruleaction-enabled-property-outlook.md)** property set to **False** . Rule actions that are not supported during programmatic rule creation can only be enumerated in the **RuleActions** collection for an existing rule, but because the **RuleActions** collection is fixed, you cannot create a rule and add such an action to the associated **RuleActions** collection.


## See also


#### Concepts


[Rule Object](rule-object-outlook.md)

