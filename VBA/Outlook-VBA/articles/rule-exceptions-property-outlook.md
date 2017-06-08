---
title: Rule.Exceptions Property (Outlook)
keywords: vbaol11.chm2176
f1_keywords:
- vbaol11.chm2176
ms.prod: outlook
api_name:
- Outlook.Rule.Exceptions
ms.assetid: 843c2690-ee39-bac7-d593-80c3dd31087f
ms.date: 06/08/2017
---


# Rule.Exceptions Property (Outlook)

Returns a  **[RuleConditions](ruleconditions-object-outlook.md)** collection object that represents all the available rule exception conditions for the rule. Read-only.


## Syntax

 _expression_ . **Exceptions**

 _expression_ A variable that represents a **Rule** object.


## Remarks

An exception condition for a rule states the condition under which the rule should not be applied. Both the  **[Conditions](rule-conditions-property-outlook.md)** and **Exceptions** properties share the same pool of conditions and return a corresponding **RuleConditions** collection object.

You can enumerate and enable rules with any rule exception condition that the Rules and Alerts Wizard support, but you can programmatically create rules that have only the most commonly used rule exception conditions, and not any rule exception condition that the Rules and Alerts Wizard supports. For more information on rule condition support, see [Specify Rule Conditions](http://msdn.microsoft.com/library/812c131a-fe23-1b8b-5e2d-9459d7102630%28Office.15%29.aspx).

Through the  **Conditions** property, each rule is associated with a **RuleConditions** object. The **RuleConditions** collection is a fixed object - you cannot add or remove items from this collection. Rule exception conditions that are enabled in the rule will have an enabled rule exception condition in the **RuleConditions** collection. Rule exception conditions that are not enabled in the rule will have a rule exception condition in this collection that has the **[RuleCondition.Enabled](rulecondition-enabled-property-outlook.md)** property set to **False** . Rule exception conditions that are not supported during programmatic rule creation can only be enumerated in the **RuleConditions** collection for an existing rule, but because the **RuleConditions** collection is fixed, you cannot create a rule and add such an exception condition to the associated **RuleConditions** collection.


## See also


#### Concepts


[Rule Object](rule-object-outlook.md)

