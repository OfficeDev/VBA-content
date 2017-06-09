---
title: OlRuleConditionType Enumeration (Outlook)
keywords: vbaol11.chm3116
f1_keywords:
- vbaol11.chm3116
ms.prod: outlook
api_name:
- Outlook.OlRuleConditionType
ms.assetid: 35c2f965-0f9d-8cc8-2f05-60522268574f
ms.date: 06/08/2017
---


# OlRuleConditionType Enumeration (Outlook)

Specifies the type of rule condition or exception condition of a rule.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **olConditionAccount**|3| **Account** is the account specified in **[AccountRuleCondition.Account](accountrulecondition-account-property-outlook.md)** .|
| **olConditionAnyCategory**|29|Message is assigned to any category.|
| **olConditionBody**|13|Body contains words specified in  **[TextRuleCondition.Text](textrulecondition-text-property-outlook.md)** .|
| **olConditionBodyOrSubject**|14|Body or subject contains words specified by  **TextRuleCondition.Text.**|
| **olConditionCategory**|18| **Category** is the category specified in **[CategoryRuleCondition.Categories](categoryrulecondition-categories-property-outlook.md)** .|
| **olConditionCc**|9|Message has my name in the  **Cc** box.|
| **olConditionDateRange**|22|Message was received between x and y, where x and y are  **Date** values.|
| **olConditionFlaggedForAction**|8|Message is flagged for the specified action.|
| **olConditionFormName**|23|Message uses the form specified in  **[FormNameRuleCondition.FormName](formnamerulecondition-formname-property-outlook.md)** .|
| **olConditionFrom**|1|Sender is in the recipient list specified in  **[ToOrFromRuleCondition.Recipients](toorfromrulecondition-recipients-property-outlook.md)** .|
| **olConditionFromAnyRssFeed**|31|Message is generated from any RSS subscription.|
| **olConditionFromRssFeed**|30|Message is generated from a specific RSS subscription.|
| **olConditionHasAttachment**|20|Message has one or more attachments.|
| **olConditionImportance**|6|Message is marked with the specified level of importance.|
| **olConditionLocalMachineOnly**|27|Rule can run only on the local machine.|
| **olConditionMeetingInviteOrUpdate**|26|Message is a meeting invitation or update.|
| **olConditionMessageHeader**|15|Message header contains words specified in  **TextRuleCondition.Text** .|
| **olConditionNotTo**|11|Message does not have my name in the  **To** box.|
| **olConditionOnlyToMe**|4|Message is sent only to me.|
| **olConditionOOF**|19|Message is an out-of-office message.|
| **olConditionOtherMachine**|28|Rule can run only on a specific machine that is not the current machine.|
| **olConditionProperty**|24|Document property is exactly, contains, or does not contain specified properties.|
| **olConditionRecipientAddress**|16|Recipient address contains words specified in  **TextRuleCondition.Text** .|
| **olConditionSenderAddress**|17|Sender address contains words specified in  **TextRuleCondition.Text** .|
| **olConditionSenderInAddressBook**|25|Sender is in the address list specified in  **[AddressRuleCondition.Address](addressrulecondition-address-property-outlook.md)** .|
| **olConditionSensitivity**|7|Message is marked with the specified level of sensitivity.|
| **olConditionSentTo**|12|Sent to recipients ( **To**,  **Cc**) are in the recipient list specified in  **ToOrFromRuleCondition.Recipients** .|
| **olConditionSizeRange**|21|Message size is between x and y in units of KB, where x and y are  **Integer** values.|
| **olConditionSubject**|2|Subject contains words specified in  **TextRuleCondition.Text** .|
| **olConditionTo**|5|My name is in the  **To** box.|
| **olConditionToOrCc**|10|Message has my name in the  **To** or **Cc** box.|
| **olConditionUnknown**|0|Unrecognized condition.|

## Remarks

The list of rule condition types in this enumeration includes all the rule conditions and exception conditions that the Rules and Alerts Wizard supports. Note that while you can programmatically enumerate all the rule conditions and exception conditions for a rule, you can programmatically create a rule with only the most commonly used conditions. For more information, see [Specifying Rule Conditions](http://msdn.microsoft.com/library/812c131a-fe23-1b8b-5e2d-9459d7102630%28Office.15%29.aspx).

For example, the rule condition types  **olConditionDateRange** and **olConditionSizeRange** are supported only for enumeration and enabling or disabling in an existing rule. You cannot programmatically create a rule with such conditions. You cannot programmatically get or set the values of x and y that represent the range.


