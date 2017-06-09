---
title: OlRuleActionType Enumeration (Outlook)
keywords: vbaol11.chm3113
f1_keywords:
- vbaol11.chm3113
ms.prod: outlook
api_name:
- Outlook.OlRuleActionType
ms.assetid: d6a39ac2-00e7-73e7-3890-ea658211eae9
ms.date: 06/08/2017
---


# OlRuleActionType Enumeration (Outlook)

Specifies the type of rule action for a rule.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **olRuleActionAssignToCategory**|2|Rule action is to assign categories to the message.|
| **olRuleActionCcMessage**|27|Rule action is to cc the message to specified recipients.|
| **olRuleActionClearCategories**|30|Rule action is to clear all the categories assigned to the message.|
| **olRuleActionCopyToFolder**|5|Rule action is to copy the message to a specified folder.|
| **olRuleActionCustomAction**|22|Rule action is to perform a custom action.|
| **olRuleActionDefer**|28|Rule action is to defer delivery of the message by the specified number of minutes.|
| **olRuleActionDelete**|3|Rule action is to delete the message.|
| **olRuleActionDeletePermanently**|4|Rule action is to permanently delete the message.|
| **olRuleActionDesktopAlert**|24|Rule action is to display a desktop alert.|
| **olRuleActionFlagClear**|13|Rule action is to clear the message flag.|
| **olRuleActionFlagColor**|12|Rule action is to flag the message with a specified colored flag.|
| **olRuleActionFlagForActionInDays**|11|Rule action is to flag the message for action in the specified number of days.|
| **olRuleActionForward**|6|Rule action is to forward the message to the specified recipients.|
| **olRuleActionForwardAsAttachment**|7|Rule action is to forward the message as an attachment to the specified recipients.|
| **olRuleActionImportance**|14|Rule action is to mark the message with the specified level of importance.|
| **olRuleActionMarkAsTask**|41|Rule action is to mark the message as a task.|
| **olRuleActionMarkRead**|19|Rule action is to mark the message as read.|
| **olRuleActionMoveToFolder**|1|Rule action is to move the message to the specified folder.|
| **olRuleActionNewItemAlert**|23|Rule action is to display the specified text in the  **New Item Alert** dialog box.|
| **olRuleActionNotifyDelivery**|26|Rule action is to request delivery notification for the message being sent.|
| **olRuleActionNotifyRead**|25|Rule action is to request read notification for the message being sent.|
| **olRuleActionPlaySound**|17|Rule action is to play a sound file.|
| **olRuleActionPrint**|16|Rule action is to print the message on the default printer.|
| **olRuleActionRedirect**|8|Rule action is to redirect the message to the specified recipients.|
| **olRuleActionRunScript**|20|Rule action is to run a script.|
| **olRuleActionSensitivity**|15|Rule action is to mark the message with the specified level of sensitivity.|
| **olRuleActionServerReply**|9|Rule action is to request the server to reply with the specified mail item.|
| **olRuleActionStartApplication**|18|Rule action is to run an .exe file.|
| **olRuleActionStop**|21|Rule action is to stop processing more rules.|
| **olRuleActionTemplate**|10|Rule action is to use the specified template (.oft) file as a form template.|
| **olRuleActionUnknown**|0|Unrecognized rule action.|

## Remarks

The list of rule action types in this enumeration includes all the rule actions that the Rules and Alerts Wizard supports. Note that while you can programmatically enumerate all the rule actions for a rule, you can programmatically create a rule with only the most commonly used rule actions. For more information, see [Specifying Rule Actions](http://msdn.microsoft.com/library/c5f83c81-0e01-38aa-5ec7-3932b4443e43%28Office.15%29.aspx).


