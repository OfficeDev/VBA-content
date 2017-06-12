---
title: AssignToCategoryRuleAction Object (Outlook)
keywords: vbaol11.chm3168
f1_keywords:
- vbaol11.chm3168
ms.prod: outlook
api_name:
- Outlook.AssignToCategoryRuleAction
ms.assetid: 402f4742-72ba-2559-4e4c-e2b8248cd7f6
ms.date: 06/08/2017
---


# AssignToCategoryRuleAction Object (Outlook)

Represents an action that assigns categories to a message.


## Remarks

 **AssignToCategoryRuleAction** is derived from the **[RuleAction](ruleaction-object-outlook.md)** object. Each rule is associated with a **[RuleActions](ruleactions-object-outlook.md)** object which has an **[AssignToCategory](ruleactions-assigntocategory-property-outlook.md)** property. The **AssignToCategory** property always returns an **[AssignToCategoryRuleAction](assigntocategoryruleaction-object-outlook.md)** object. If the rule has an enabled rule action that assigns a message with some specified categories, then **[AssignToCategoryRuleAction.Enabled](assigntocategoryruleaction-enabled-property-outlook.md)** would be **True**.

For more information on specifying rule actions, see [Specify Rule Actions](http://msdn.microsoft.com/library/c5f83c81-0e01-38aa-5ec7-3932b4443e43%28Office.15%29.aspx).


## Properties



|**Name**|
|:-----|
|[ActionType](assigntocategoryruleaction-actiontype-property-outlook.md)|
|[Application](assigntocategoryruleaction-application-property-outlook.md)|
|[Categories](assigntocategoryruleaction-categories-property-outlook.md)|
|[Class](assigntocategoryruleaction-class-property-outlook.md)|
|[Enabled](assigntocategoryruleaction-enabled-property-outlook.md)|
|[Parent](assigntocategoryruleaction-parent-property-outlook.md)|
|[Session](assigntocategoryruleaction-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
