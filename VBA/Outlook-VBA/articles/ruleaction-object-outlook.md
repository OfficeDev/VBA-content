---
title: RuleAction Object (Outlook)
keywords: vbaol11.chm3163
f1_keywords:
- vbaol11.chm3163
ms.prod: outlook
api_name:
- Outlook.RuleAction
ms.assetid: 6451788f-e5ed-239c-a34d-b564b52d8955
ms.date: 06/08/2017
---


# RuleAction Object (Outlook)

Represents an action that is run when a  **[Rule](rule-object-outlook.md)** object executes.


## Remarks

 **RuleAction** is the base class for rule actions that are supported in programmatic rule creation. The classes derived from **RuleAction** include:


-  **[AssignToCategoryRuleAction](assigntocategoryruleaction-object-outlook.md)**
    
-  **[MarkAsTaskRuleAction](markastaskruleaction-object-outlook.md)**
    
-  **[MoveOrCopyRuleAction](moveorcopyruleaction-object-outlook.md)**
    
-  **[NewItemAlertRuleAction](newitemalertruleaction-object-outlook.md)**
    
-  **[PlaySoundRuleAction](playsoundruleaction-object-outlook.md)**
    
-  **[SendRuleAction](sendruleaction-object-outlook.md)**
    


The Rules object model provides partial parity with the Rules and Alerts Wizard in the Outlook user interface. It supports the most commonly used rule actions and conditions. Although it does not support creating rules with each rule action or rule condition that the Wizard supports, you can still enumerate and enable these rule actions and conditions in existing rules. 

For more information on rule actions, see [Specifying Rule Actions](http://msdn.microsoft.com/library/c5f83c81-0e01-38aa-5ec7-3932b4443e43%28Office.15%29.aspx) and[How to: Create a Rule to Move Specific E-mails to a Folder](http://msdn.microsoft.com/library/e72fa307-8224-c2d2-1318-a18cd8e9f22f%28Office.15%29.aspx).


## Properties



|**Name**|
|:-----|
|[ActionType](ruleaction-actiontype-property-outlook.md)|
|[Application](ruleaction-application-property-outlook.md)|
|[Class](ruleaction-class-property-outlook.md)|
|[Enabled](ruleaction-enabled-property-outlook.md)|
|[Parent](ruleaction-parent-property-outlook.md)|
|[Session](ruleaction-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
