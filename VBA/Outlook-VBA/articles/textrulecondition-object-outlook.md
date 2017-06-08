---
title: TextRuleCondition Object (Outlook)
keywords: vbaol11.chm3183
f1_keywords:
- vbaol11.chm3183
ms.prod: outlook
api_name:
- Outlook.TextRuleCondition
ms.assetid: 87e9ca00-7577-02c2-fb6f-a5dc2054ad8b
ms.date: 06/08/2017
---


# TextRuleCondition Object (Outlook)

Represents a rule condition that the part of the message, which can be the body, header, or subject, as specified by  **[TextRuleCondition.ConditionType](textrulecondition-conditiontype-property-outlook.md)**, contains the words specified in **[TextRuleCondition.Text](textrulecondition-text-property-outlook.md)**.


## Remarks

 **TextRuleCondition** is derived from the **[RuleCondition](rulecondition-object-outlook.md)** object. Each rule is associated with a **[RuleConditions](ruleconditions-object-outlook.md)** object which has the following properties: **[Body](ruleconditions-body-property-outlook.md)**, **[BodyOrSubject](ruleconditions-bodyorsubject-property-outlook.md)**, **[MessageHeader](ruleconditions-messageheader-property-outlook.md)**, and **[Subject](ruleconditions-subject-property-outlook.md)**. Each of these properties always returns a **TextRuleCondition** object. **TextRuleCondition.ConditionType** distinguishes among these rule conditions. If the rule has any of these rule conditions enabled, then **[TextRuleCondition.Enabled](textrulecondition-enabled-property-outlook.md)** would be **True**.

For more information on specifying rule conditions, see [Specify Rule Conditions](http://msdn.microsoft.com/library/812c131a-fe23-1b8b-5e2d-9459d7102630%28Office.15%29.aspx).


## Properties



|**Name**|
|:-----|
|[Application](textrulecondition-application-property-outlook.md)|
|[Class](textrulecondition-class-property-outlook.md)|
|[ConditionType](textrulecondition-conditiontype-property-outlook.md)|
|[Enabled](textrulecondition-enabled-property-outlook.md)|
|[Parent](textrulecondition-parent-property-outlook.md)|
|[Session](textrulecondition-session-property-outlook.md)|
|[Text](textrulecondition-text-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
