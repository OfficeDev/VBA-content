---
title: Specifying Rule Conditions
ms.prod: outlook
ms.assetid: 812c131a-fe23-1b8b-5e2d-9459d7102630
ms.date: 06/08/2017
---


# Specifying Rule Conditions

The Rules object model supports the most commonly used rule actions and conditions. Each  **[Rule](rule-object-outlook.md)** object has an **[Actions](rule-actions-property-outlook.md)** property that represents the rule actions for that rule, as well as a **[Conditions](rule-conditions-property-outlook.md)** property and an **[Exceptions](rule-exceptions-property-outlook.md)** property that represent the conditions for that rule. This topic describes how the Rules object model supports rule conditions.

Rule conditions for a rule are represented by a  **[RuleConditions](ruleconditions-object-outlook.md)** collection object. A **RuleConditions** object has properties that correspond to each commonly used rule condition in a rule. For example, if a rule specifies two conditions - the message is important and the subject contains certain words - then the **[Importance](ruleconditions-importance-property-outlook.md)** and **[Subject](ruleconditions-subject-property-outlook.md)** properties of the rule's **RuleConditions** collection object will return respective rule condition objects that are enabled (**[RuleCondition.Enabled](rulecondition-enabled-property-outlook.md)** is **True**). 

Conditions that are not specified in a rule will not be enabled in the corresponding  **[RuleCondition](rulecondition-object-outlook.md)** object (**RuleCondition.Enabled** is **False**). Rule condition objects are represented by either the  **RuleCondition** object or customized objects derived from the **RuleCondition** object. In the last example, the **RuleConditions.Importance** property will return an **[ImportanceRuleCondition](importancerulecondition-object-outlook.md)** object, and the **RuleConditions.Subject** property will return a **[TextRuleCondition](textrulecondition-object-outlook.md)** object, both of which are derived from the **RuleCondition** object. The **RuleCondition** object and its derived objects have the **ConditionType** property that will indicate the type of the rule condition, for example, **[ImportanceRuleCondition.ConditionType](importancerulecondition-conditiontype-property-outlook.md)** will indicate the value **olConditionImportance**, and  **[TextRuleCondition.ConditionType](textrulecondition-conditiontype-property-outlook.md)** will indicate **olConditionSubject**. 

Note that the Rules object model maintains partial parity with the Rules and Alerts Wizard. This means that while you can use the Wizard to create rules that specify any action and condition that you see in the Wizard, you can programmatically create rules that use some but not all of these actions and conditions. An example of a condition that the Rules object model supports for rules created by the Wizard but not for those created by the object model is messages of certain level of sensitivity. You can use the Wizard to create a rule specifying sensitivity as a condition. 

Using the Rules object model, you can enumerate this kind of rule in the  **Rules** collection - for each rule in the **Rules** collection, enumerate its **RuleConditions** collection and look for an enabled rule condition for sensitivity. In code, this would mean for each rule in the **Rules** collection, enumerate **[RuleConditions.Item(Index)](ruleconditions-item-method-outlook.md)** using the _Index_ from 1 to **[RuleConditions.Count](ruleconditions-count-property-outlook.md)** and look for an enabled condition with **[RuleCondition.ConditionType](rulecondition-conditiontype-property-outlook.md)** equal to **olConditionSensitivity**. You can also enable or disable such a rule condition in a rule. However, you cannot programmatically create a rule that specifies the  **olConditionSensitivity** condition.

The following table lists all the rule conditions supported by the Rules and Alerts Wizard, and whether each rule condition is supported when creating a rule using the Rules object model. A rule condition that is not supported in rules created by the Rules object model is supported only for programmatic enumeration and enabling or disabling in existing rules created by the Rules and Alerts Wizard. The table also shows whether the rule condition applies to rules with the  **olRuleReceive** rule type or **olRuleSend** rule type, or both.

 **Note**  You cannot enable or disable a rule condition of the type  **olConditionOtherMachine**. This type of rule condition indicates that the rule can run only on a certain computer, but the current computer is not that computer. This happens when the rule is created on one computer and the rule condition  **olConditionLocalMachineOnly** is enabled, indicating that the rule can run only on that computer. In certain cases, **olConditionLocalMachine** is automatically set as a result of enabling another rule condition such as **olConditionAccount**. When you run the same rule on another computer, the rule will show that the condition  **olConditionOtherMachine** is enabled.



| **Condition**| **Constant in olRuleConditionType**| **Supported when creating new rules programmatically?**| **Apply to olRuleReceive rules?**| **Apply to olRuleSend rules?**|
|:-----|:-----|:-----|:-----|:-----|
|Account is the account specified in  **[AccountRuleCondition.Account](accountrulecondition-account-property-outlook.md)**.| **olConditionAccount**|Yes|Yes|Yes|
|Message is assigned any category.| **olCondtionAnyCategory**|Yes|Yes|Yes|
|Body contains words specified in  **[TextRuleCondition.Text](textrulecondition-text-property-outlook.md)**.| **olConditionBody**|Yes|Yes|Yes|
|Body or subject contains words specified by  **TextRuleCondition.Text.**| **olConditionBodyOrSubject**|Yes|Yes|Yes|
|Message is assigned the category or categories specified in  **[CategoryRuleCondition.Categories](categoryrulecondition-categories-property-outlook.md)**.| **olConditionCategory**|Yes|Yes|Yes|
|Message has my name in the  **Cc** box.| **olConditionCc**|Yes|Yes||
|Message was received between x and y, where x and y are Integer values. | **olConditionDateRange**|No|Yes|Yes|
|Message is flagged for the specified action.| **olConditionFlaggedForAction**|No|Yes|Yes|
|Message uses the form specified in  **[FormNameRuleCondition.FormName](formnamerulecondition-formname-property-outlook.md)**.| **olConditionFormName**|Yes|Yes|Yes|
|Sender is in the recipient list specified in  **[ToOrFromRuleCondition.Recipients](toorfromrulecondition-recipients-property-outlook.md)**.| **olConditionFrom**|Yes|Yes|No|
|Message is generated from any RSS subscription.| **olCondtionFromAnyRssFeed**|Yes|Yes|No|
|Message is generated from a specified RSS subscription.| **olCondtionFromRssFeed**|Yes|Yes|No|
|Message has an attachment.| **olConditionHasAttachment**|Yes|Yes|Yes|
|Message is marked with the specified level of importance.| **olConditionImportance**|Yes|Yes|Yes|
|Rule can run only on this machine.| **olConditionLocalMachineOnly**|Yes|Yes|Yes|
|Message is a meeting invitation or update.| **olConditionMeetingInviteOrUpdate**|Yes|Yes|Yes|
|Message header contains words specified in  **TextRuleCondition.Text**.| **olConditionMessageHeader**|Yes|Yes|No|
|Message does not have my name in the  **To** box.| **olConditionNotTo**|Yes|Yes|No|
|Message is sent only to me.| **olConditionOnlyToMe**|Yes|Yes|No|
|Message is an out-of-office message.| **olConditionOOF**|No|Yes|No|
|Rule can run only on a specific machine that is not the current one.| **olConditionOtherMachine**|No|Yes|Yes|
|Document property is exactly, contains, or does not contain specified properties.| **olConditionProperty**|No|Yes|Yes|
|Recipient address contains words specified in  **TextRuleCondition.Text**.| **olConditionRecipientAddress**|Yes|Yes|Yes|
|Sender address contains words specified in  **TextRuleCondition.Text**.| **olConditionSenderAddress**|Yes|Yes|No|
|Sender is in the address list specified in  **[AddressRuleCondition.Address](addressrulecondition-address-property-outlook.md)**.| **olConditionSenderInAddressBook**|Yes|Yes|No|
|Message is marked with the specified level of sensitivity.| **olConditionSensitivity**|No|Yes|Yes|
|Sent to recipients ( **To**,  **Cc**) are in the recipient list specified in  **ToOrFromRuleCondition.Recipients**.| **olConditionSentTo**|Yes|Yes|Yes|
|Message size is between x and y in units of KB, where x and y are  **Date** values. For example, "10;50" sets the size condition between 10 and 50KB.| **olConditionSizeRange**|No|Yes|Yes|
|Subject contains words specified in  **TextRuleCondition.Text**.| **olConditionSubject**|Yes|Yes|Yes|
|My name is in the  **To** box.| **olConditionTo**|Yes|Yes|No|
|Message has my name in the  **To** or **Cc** box.| **olConditionToOrCc**|Yes|Yes|No|
|Unrecognized rule condition.| **olConditionUnknown**|No|Yes|No|


