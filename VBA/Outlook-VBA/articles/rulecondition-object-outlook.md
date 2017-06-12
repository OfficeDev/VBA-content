---
title: RuleCondition Object (Outlook)
keywords: vbaol11.chm3173
f1_keywords:
- vbaol11.chm3173
ms.prod: outlook
api_name:
- Outlook.RuleCondition
ms.assetid: e03f91c2-2c08-b036-104a-d6246f28bc2d
ms.date: 06/08/2017
---


# RuleCondition Object (Outlook)

The  **RuleCondition** object represents either a condition that must be met before a rule executes, or an exception condition that must not be met before a rule executes.


## Remarks

 **RuleCondition** is the base class for rule conditions that are supported in programmatic rule creation. The classes derived from **RuleCondition** include:


-  **[AccountRuleCondition](accountrulecondition-object-outlook.md)**
    
-  **[AddressRuleCondition](addressrulecondition-object-outlook.md)**
    
-  **[CategoryRuleCondition](categoryrulecondition-object-outlook.md)**
    
-  **[FromRssFeedRuleCondition](fromrssfeedrulecondition-object-outlook.md)**
    
-  **[FormNameRuleCondition](formnamerulecondition-object-outlook.md)**
    
-  **[ImportanceRuleCondition](importancerulecondition-object-outlook.md)**
    
-  **[SenderInAddressListRuleCondition](senderinaddresslistrulecondition-object-outlook.md)**
    
-  **[TextRuleCondition](textrulecondition-object-outlook.md)**
    
-  **[ToOrFromRuleCondition](toorfromrulecondition-object-outlook.md)**
    


The Rules object model provides partial parity with the Rules and Alerts Wizard in the Outlook user interface. It supports the most commonly used rule actions and conditions. Although it does not support creating rules with each rule action or rule condition that the Wizard supports, you can still enumerate and enable these rule actions and conditions in existing rules. 

For more information on rule conditions, see [Specifying Rule Conditions](http://msdn.microsoft.com/library/812c131a-fe23-1b8b-5e2d-9459d7102630%28Office.15%29.aspx) and[How to: Create a Rule to Move Specific E-mails to a Folder](http://msdn.microsoft.com/library/e72fa307-8224-c2d2-1318-a18cd8e9f22f%28Office.15%29.aspx).


## Properties



|**Name**|
|:-----|
|[Application](rulecondition-application-property-outlook.md)|
|[Class](rulecondition-class-property-outlook.md)|
|[ConditionType](rulecondition-conditiontype-property-outlook.md)|
|[Enabled](rulecondition-enabled-property-outlook.md)|
|[Parent](rulecondition-parent-property-outlook.md)|
|[Session](rulecondition-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
