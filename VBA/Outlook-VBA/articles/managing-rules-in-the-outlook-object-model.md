---
title: Managing Rules in the Outlook Object Model
ms.prod: outlook
ms.assetid: 05ddd643-e9bd-a37d-b680-b8519960a5f6
ms.date: 06/08/2017
---


# Managing Rules in the Outlook Object Model

The  **Rules** object model supports the programmatic adding, editing, and deleting of rules. The **[Rule](rule-object-outlook.md)** and **[Rules](rules-object-outlook.md)** collection objects allow you to access, add, and delete rules defined for a session. The **[RuleAction](ruleaction-object-outlook.md)** and **[RuleCondition](rulecondition-object-outlook.md)** objects, their collection objects, and derived action and condition objects further support editing actions and conditions.


 **Note**  The  **Rules** object model provides partial parity with the **Rules and Alerts Wizard** in the Outlook user interface. Although it does not support every single rule that you can possibly create using the Wizard, it supports the most commonly used rule actions and conditions. Just like any rule created by using the **Rules and Alerts** Wizard, rules created programmatically are applied to messages, which include mail items, meeting requests, task requests, documents, delivery receipts, read receipts, voting responses, and out-of-office notices.


Use  **[Store.GetRules](store-getrules-method-outlook.md)** to obtain a **Rules** collection object representing the rules defined for the store used in the current session.

After obtaining the set of rules for the current session, you can then add new rules (by using  **[Rules.Create](rules-create-method-outlook.md)**), edit existing rules (by enabling or disabling rules, changing their execution order, and modifying rule actions and rule conditions), or delete rules (by using  **[Rules.Remove](rules-remove-method-outlook.md)**) from this  **Rules** collection. Note that while you can edit rules created in versions of Outlook before Microsoft Office Outlook 2007, you cannot use earlier versions of Outlook to edit rules that have been created in Office Outlook 2007 or later.

You can retrieve each rule in a  **Rules** collection by indexing the collection using **[Rules.Item(Index)](rules-item-method-outlook.md)**, with  _Index_ being either the name of the rule (the default property **[Rule.Name](rule-name-property-outlook.md)**), or a value ranging from 1 through the total number of rules in the collection,  **[Rules.Count](rules-count-property-outlook.md)**. 

**[Rule.ExecutionOrder](rule-executionorder-property-outlook.md)** indicates the order of execution of the rules in the collection and is directly mapped with the numerical value of _Index_ in **Rules.Items(Index)**. For example,  `Rules.Item(1)` represents a rule with **Rule.ExecutionOrder** being 1, `Rules.Item(2)` represents a rule with **Rule.ExecutionOrder** being 2, and `Rules.Item(Rules.Count)` represents the rule with **Rule.ExecutionOrder** being **Rules.Count**.

After you have defined a rule, you should also enable it by setting the  **[Rule.Enabled](rule-enabled-property-outlook.md)** property to **True**, and then save the rule by using  **[Rules.Save](rules-save-method-outlook.md)** so that the rule and its enabled state will persist beyond the current session. A rule is only enabled after it has been saved successfully. Note that **Rules.Save** can be an expensive operation in terms of performance on slow connections to Exchange server; you can optionally display a progress dialog box for the user to cancel the operation. A save fails when the user edits the same rule in the Rules and Alerts Wizard, or the user cancels the progress dialog box. In such cases, **Rules.Save** will raise an error, and the user will resolve the conflict by responding to the error dialog brought up by the Rules and Alerts Wizard.

When you use  **Rules.Save** to save one or more rules that have been created in Office Outlook 2007, you will be prompted with a dialog to remind you that you will not be able to edit that rule using earlier versions of Outlook. You will have to confirm the dialog before the save opreation can proceed.

Use  **[Rule.Execute](rule-execute-method-outlook.md)** to run a rule. Note that while you must enable and save a rule to have it enabled beyond the current session, you can run the rule regardless of its enabled state. When you execute a rule, you can optionally specify the folder to apply to rule to. The default is to execute the rule against all messages in the Inbox, but not subfolders of the Inbox.

