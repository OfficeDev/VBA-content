---
title: TextRuleCondition.ConditionType Property (Outlook)
keywords: vbaol11.chm2477
f1_keywords:
- vbaol11.chm2477
ms.prod: outlook
api_name:
- Outlook.TextRuleCondition.ConditionType
ms.assetid: 2dbc7979-deae-fbb8-9def-8c906657024a
ms.date: 06/08/2017
---


# TextRuleCondition.ConditionType Property (Outlook)

Returns a constant from the  **[OlRuleConditionType](olruleconditiontype-enumeration-outlook.md)** enumeration that indicates the type of rule condition. Read-only.


## Syntax

 _expression_ . **ConditionType**

 _expression_ A variable that represents a **TextRuleCondition** object.


## Remarks

The value of  **ConditionType** depends on the type of rule condition, as several types of rule conditions use the **[TextRuleCondition](textrulecondition-object-outlook.md)** object: **olConditionBody** , **olConditionBodyOrSubject** , **olConditionMessageHeader** , and **olConditionSubject** . Except for **olConditionMessageHeader** , which is supported only by rules for receiving messages, all these types of conditions are supported by rules for receiving messages as well as rules for sending messages. For more information, see[Specify Rule Conditions](http://msdn.microsoft.com/library/812c131a-fe23-1b8b-5e2d-9459d7102630%28Office.15%29.aspx).


## See also


#### Concepts


[TextRuleCondition Object](textrulecondition-object-outlook.md)

