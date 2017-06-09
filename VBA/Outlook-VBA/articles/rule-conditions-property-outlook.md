---
title: Rule.Conditions Property (Outlook)
keywords: vbaol11.chm2175
f1_keywords:
- vbaol11.chm2175
ms.prod: outlook
api_name:
- Outlook.Rule.Conditions
ms.assetid: e2cacf1c-95eb-31d3-012c-7cf9426053d5
ms.date: 06/08/2017
---


# Rule.Conditions Property (Outlook)

Returns a  **[RuleConditions](ruleconditions-object-outlook.md)** collection object that represents all the available rule conditions for the rule. Read-only.


## Syntax

 _expression_ . **Conditions**

 _expression_ A variable that represents a **Rule** object.


## Remarks

A condition for a rule states the condition under which the rule should be applied. Both the  **Conditions** and **[Exceptions](rule-exceptions-property-outlook.md)** properties share the same pool of conditions and return a corresponding **RuleConditions** collection object.

Programmatically you can enumerate and enable rules with any rule condition that the Rules and Alerts Wizard support, but you can create rules that have only the most commonly used rule conditions, and not any rule condition that the Rules and Alerts Wizard supports. For more information on rule condition support, see [Specify Rule Conditions](http://msdn.microsoft.com/library/812c131a-fe23-1b8b-5e2d-9459d7102630%28Office.15%29.aspx).

Through the  **Conditions** property, each rule is associated with a **RuleConditions** object. The **RuleConditions** collection is a fixed object - you cannot add or remove items from this collection. Rule conditions that are enabled in the rule will have an enabled rule condition in the **RuleConditions** collection. Rule conditions that are not enabled in the rule will have a rule condition in this collection that has the **[RuleCondition.Enabled](rulecondition-enabled-property-outlook.md)** property set to **False** . Rule conditions that are not supported during programmatic rule creation can only be enumerated in the **RuleConditions** collection for an existing rule, but because the **RuleConditions** collection is fixed, you cannot create a rule and add such a condition to the associated **RuleConditions** collection.


## See also


#### Concepts


[Rule Object](rule-object-outlook.md)

