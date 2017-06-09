---
title: Specifying Rule Actions
ms.prod: outlook
ms.assetid: c5f83c81-0e01-38aa-5ec7-3932b4443e43
ms.date: 06/08/2017
---


# Specifying Rule Actions

The Rules object model supports the most commonly used rule actions and conditions. Each  **[Rule](rule-object-outlook.md)** object has an **[Actions](rule-actions-property-outlook.md)** property that represents the rule actions for that rule, as well as a **[Conditions](rule-conditions-property-outlook.md)** property and an **[Exceptions](rule-exceptions-property-outlook.md)** property that represent the conditions for that rule. This topic describes how the Rules object model supports rule actions.

Rule actions for a rule are represented by a  **[RuleActions](ruleactions-object-outlook.md)** collection object. A **RuleActions** object has properties that correspond to each commonly used rule action in a rule. For example, if a rule specifies two actions - moving the message to a specific folder and plays a sound - then the **[MoveToFolder](ruleactions-movetofolder-property-outlook.md)** and **[PlaySound](ruleactions-playsound-property-outlook.md)** properties of the rule's **RuleActions** collection object will return respective rule action objects that are enabled ( **[RuleAction.Enabled](ruleaction-enabled-property-outlook.md)** is **True**). 

Actions that are not specified in a rule will not be enabled in the corresponding  **RuleAction** object (**RuleAction.Enabled** is **False**). These rule action objects are represented by either the  **RuleAction** object or customized objects derived from the **RuleAction** object. In the last example, specifically, the **RuleActions.MoveToFolder** property will return a **[MoveOrCopyRuleAction](moveorcopyruleaction-object-outlook.md)** object, and the **RuleActions.PlaySound** property will return a **[PlaySoundRuleAction](playsoundruleaction-object-outlook.md)** object, both of which are derived from the **RuleAction** object. The **RuleAction** object and its derived objects have the **ActionType** property that will indicate the type of the rule action. For example, **[MoveOrCopyRuleAction.ActionType](moveorcopyruleaction-actiontype-property-outlook.md)** will indicate the value **olRuleActionMoveToFolder**, and  **[PlaySoundRuleAction.ActionType](playsoundruleaction-actiontype-property-outlook.md)** will indicate **olRuleActionPlay**. 

Note that the Rules object model maintains partial parity with the Rules and Alerts Wizard. This means that while you can use the Wizard to create rules that specify any action and condition that you see in the Wizard, you can programmatically create rules that use some but not all of these actions and conditions. An example of an action that the Rules object model supports for rules created by the Wizard but not for those created by the object model is requesting a server reply. You can use the Wizard to create a rule specifying a certain server reply as an action. 

Using the Rules object model, you can enumerate these kinds of rules in the  **Rules** collection - for each rule in the **Rules** collection, enumerate its **RuleActions** collection and look for an enabled rule action for a server reply. In code, this would mean for each rule in the **Rules** collection, enumerate **[RuleActions.Item(Index)](ruleactions-item-method-outlook.md)** using the _Index_ from 1 to **[RuleActions.Count](ruleactions-count-property-outlook.md)**, and look for an enabled action with  **ActionType** equal to **olRuleActionServerReply**. You can also enable or disable such a rule action in a rule. However, you cannot programmatically create a rule that specifies the  **olRuleActionServerReply** action.

The following table lists all the rule actions supported by the Rules and Alerts Wizard, and whether each rule action is supported when creating a rule using the Rules object model. A rule action that is not supported in rules created by the Rules object model is supported only for programmatic enumeration and enabling or disabling in existing rules created by the Rules and Alerts Wizard. The table also shows whether the rule action applies to rules with the  **olRuleReceive** rule type or **olRuleSend** rule type, or both.


| **Action**| **Constant in olRuleActionType**| **Supported when creating new rules programmatically?**| **Apply to olRuleReceive rules?**| **Apply to olRuleSend rules?**|
|:-----|:-----|:-----|:-----|:-----|
|Assign the message to the categories specified in the  **[AssignToCategoryRuleAction.Categories](assigntocategoryruleaction-categories-property-outlook.md)** property| **olRuleActionAssignToCategory**|Yes|Yes|Yes|
|Cc the message to the recipient list specified in the  **[SendRuleAction.Recipients](sendruleaction-recipients-property-outlook.md)** property| **olRuleActionCcMessage**|Yes|No|Yes|
|Clear all categories for the message.| **olRuleActionClearCategories**|Yes|Yes|Yes|
|Copy the message to folder specified in the **[MoveOrCopyRuleAction.Folder](moveorcopyruleaction-folder-property-outlook.md)** property| **olRuleActionCopyToFolder**|Yes|Yes|Yes|
|Run a custom action| **olRuleActionCustomAction**|No|Yes|Yes|
|Defer the delivery by a specified number of minutes| **olRuleActionDefer**|No|No|Yes|
|Delete the message| **olRuleActionDelete**|Yes|Yes|No|
|Permanently delete the message| **olRuleActionDeletePermanently**|Yes|Yes|No|
|Display a desktop alert| **olRuleActionDesktopAlert**|Yes|Yes|No|
|Clear the message flag| **olRuleActionFlagClear**|No|Yes|No|
|Flag the message with the color specified | **olRuleActionFlagColor**|No|Yes|No|
|Flag the message for action in days specified | **olRuleActionFlagForActionInDays**|No|Yes|Yes|
|Forward the message to the recipient list specified in the  **SendRuleAction.Recipients** property| **olRuleActionForward**|Yes|Yes|No|
|Forward the message as an attachment to the recipient list specified in the  **SendRuleAction.Recipients** property| **olRuleActionForwardAsAttachment**|Yes|Yes|No|
|Mark the message with the specified Importance| **olRuleActionImportance**|No|Yes|Yes|
|Mark message as a task for followup using the  **[FlagTo](markastaskruleaction-flagto-property-outlook.md)** and **[MarkInterval](markastaskruleaction-markinterval-property-outlook.md)** properties of the **[MarkAsTaskRuleAction](markastaskruleaction-object-outlook.md)** object| **olRuleActionMarkAsTask**|Yes|Yes|No|
|Mark as read| **olRuleActionMarkRead**|No|Yes|No|
|Move the message to the folder specified in the  **MoveOrCopyRuleAction.Folder** property| **olRuleActionMoveToFolder**|Yes|Yes|No|
|Display the message specified in the  **[NewItemAlertRuleAction.Text](newitemalertruleaction-text-property-outlook.md)** property| **olRuleActionNewItemAlert**|Yes|Yes|No|
|Notify that the message has been delivered| **olRuleActionNotifyDelivery**|Yes|No|Yes|
|Notify that the message has been read| **olRuleActionNotifyRead**|Yes|No|Yes|
|Play the .wav file specified in the  **[PlaySoundRuleAction.FilePath](playsoundruleaction-filepath-property-outlook.md)** property| **olRuleActionPlaysound**|Yes|Yes|No|
|Print the message to the default printer| **olRuleActionPrint**|No|Yes|No|
|Redirect the message to the recipient list specified in the  **SendRuleAction.Recipients** property| **olRuleActionRedirect**|Yes|Yes|No|
|Start a script| **olRuleActionRunScript**|No|Yes|No|
|Mark the message with the specified sensitivity| **olRuleActionSensitivity**|No|No|Yes|
|Have server reply using the specified message | **olRuleActionServerReply**|No|Yes|No|
|Start an .exe| **olRuleActionStartApplication**|No|Yes|No|
|Stop processing more rules| **olRuleActionStop**|Yes|Yes|Yes|
|Reply using the specified template (.oft) file| **olRuleActionTemplate**|No|Yes|No|
|Unrecognized rule action| **olRuleActionUnknown**|No|Yes|No|


