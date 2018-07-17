---
title: SendRuleAction Object (Outlook)
keywords: vbaol11.chm3165
f1_keywords:
- vbaol11.chm3165
ms.prod: outlook
api_name:
- Outlook.SendRuleAction
ms.assetid: 4ea8f519-8bb3-b0bf-9742-8a492e7ffff7
ms.date: 06/08/2017
---


# SendRuleAction Object (Outlook)

Represents an action that sends a message to one or more recipients.


## Remarks

 **SendRuleAction** is derived from the **[RuleAction](ruleaction-object-outlook.md)** object. Each rule is associated with a **[RuleActions](ruleactions-object-outlook.md)** object which has a **[CC](ruleactions-cc-property-outlook.md)** property, a **[Forward](ruleactions-forward-property-outlook.md)** property, a **[ForwardAsAttachment](ruleactions-forwardasattachment-property-outlook.md)** property, and a **[Redirect](ruleactions-redirect-property-outlook.md)** property. Each of these properties always returns a **SendRuleAction** object. **[SendRuleAction.ActionType](sendruleaction-actiontype-property-outlook.md)** distinguishes among these rule actions. If the rule has any of the above rule actions enabled, then the **[Enabled](sendruleaction-enabled-property-outlook.md)** property of the corresponding **SendRuleAction** object would be **True**.

For more information on specifying rule actions, see [Specify Rule Actions](http://msdn.microsoft.com/library/c5f83c81-0e01-38aa-5ec7-3932b4443e43%28Office.15%29.aspx).


## Properties



|**Name**|
|:-----|
|[ActionType](sendruleaction-actiontype-property-outlook.md)|
|[Application](sendruleaction-application-property-outlook.md)|
|[Class](sendruleaction-class-property-outlook.md)|
|[Enabled](sendruleaction-enabled-property-outlook.md)|
|[Parent](sendruleaction-parent-property-outlook.md)|
|[Recipients](sendruleaction-recipients-property-outlook.md)|
|[Session](sendruleaction-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
