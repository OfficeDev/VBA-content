---
title: RuleActions.PlaySound Property (Outlook)
keywords: vbaol11.chm2197
f1_keywords:
- vbaol11.chm2197
ms.prod: outlook
api_name:
- Outlook.RuleActions.PlaySound
ms.assetid: 43a79f2d-9e7b-7053-6901-40e815220ac0
ms.date: 06/08/2017
---


# RuleActions.PlaySound Property (Outlook)

Returns a  **[PlaySoundRuleAction](playsoundruleaction-object-outlook.md)** object with **[PlaySoundRuleAction.ActionType](playsoundruleaction-actiontype-property-outlook.md)** being **olRuleActionNotifyRead** . Read-only.


## Syntax

 _expression_ . **PlaySound**

 _expression_ A variable that represents a **RuleActions** object.


## Remarks

Use the returned  **PlaySoundRuleAction** object when enumerating the rule actions of an existing rule or when creating a new rule that specifies playing a sound file as an action.

This property of the  **[RuleActions](ruleactions-object-outlook.md)** collection always returns a **PlaySoundRuleAction** object regardless of whether the rule associated with this **RuleActions** collection has defined such a rule action. If the rule has defined and enabled such a rule action, then **[PlaySoundRuleAction.Enabled](playsoundruleaction-enabled-property-outlook.md)** will be **True** .


## See also


#### Concepts


[RuleActions Object](ruleactions-object-outlook.md)

