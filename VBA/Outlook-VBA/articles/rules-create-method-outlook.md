---
title: Rules.Create Method (Outlook)
keywords: vbaol11.chm2160
f1_keywords:
- vbaol11.chm2160
ms.prod: outlook
api_name:
- Outlook.Rules.Create
ms.assetid: 84789ccc-a6c2-9f79-5338-45b03b116dd5
ms.date: 06/08/2017
---


# Rules.Create Method (Outlook)

Creates a  **[Rule](rule-object-outlook.md)** object with the name specified by _Name_ and the type of rule specified by _RuleType_ .


## Syntax

 _expression_ . **Create**( **_Name_** , **_RuleType_** )

 _expression_ A variable that represents a **Rules** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|A string identifier for the rule, which will be represented by  **[Rule.Name](rule-name-property-outlook.md)** after rule creation. Names of rules in a collection are not unique.|
| _RuleType_|Required| **[OlRuleType](olruletype-enumeration-outlook.md)**|A constant in the  **OlRuleType** enumeration that determines whether the rule is applied on sending or receiving a message.|

### Return Value

A  **Rule** object that represents the newly created rule.


## Remarks

The  _RuleType_ parameter of the added rule determines valid rule actions, rule conditions, and rule exception conditions that can be associated with the **Rule** object.

When a rule is added to the collection, the  **[Rule.ExecutionOrder](rule-executionorder-property-outlook.md)** of the new rule is 1. The **ExecutionOrder** of other rules in the collection is incremented by 1.


## Example

The following code sample in Visual Basic for Applicatons (VBA) uses the Rules object model to create a rule. The code sample uses the  **[RuleAction](ruleaction-object-outlook.md)** and **[RuleCondition](rulecondition-object-outlook.md)** objects to specify a rule that forwards messages from a specific sender to a specific folder, unless the message contains certain terms in the subject. Note that the code sample assumes that there already exists a folder "Dan" under the Inbox.


```vb
Sub CreateRule() 
 
 Dim colRules As Outlook.Rules 
 
 Dim oRule As Outlook.Rule 
 
 Dim colRuleActions As Outlook.RuleActions 
 
 Dim oMoveRuleAction As Outlook.MoveOrCopyRuleAction 
 
 Dim oFromCondition As Outlook.ToOrFromRuleCondition 
 
 Dim oExceptSubject As Outlook.TextRuleCondition 
 
 Dim oInbox As Outlook.Folder 
 
 Dim oMoveTarget As Outlook.Folder 
 
 
 
 'Specify target folder for rule move action 
 
 Set oInbox = Application.Session.GetDefaultFolder(olFolderInbox) 
 
 'Assume that target folder already exists 
 
 Set oMoveTarget = oInbox.Folders("Dan") 
 
 
 
 'Get Rules from Session.DefaultStore object 
 
 Set colRules = Application.Session.DefaultStore.GetRules() 
 
 
 
 'Create the rule by adding a Receive Rule to Rules collection 
 
 Set oRule = colRules.Create("Dan's rule", olRuleReceive) 
 
 
 
 'Specify the condition in a ToOrFromRuleCondition object 
 
 'Condition is if the message is sent by "DanWilson" 
 
 Set oFromCondition = oRule.Conditions.From 
 
 With oFromCondition 
 
 .Enabled = True 
 
 .Recipients.Add ("DanWilson") 
 
 .Recipients.ResolveAll 
 
 End With 
 
 
 
 'Specify the action in a MoveOrCopyRuleAction object 
 
 'Action is to move the message to the target folder 
 
 Set oMoveRuleAction = oRule.Actions.MoveToFolder 
 
 With oMoveRuleAction 
 
 .Enabled = True 
 
 .Folder = oMoveTarget 
 
 End With 
 
 
 
 'Specify the exception condition for the subject in a TextRuleCondition object 
 
 'Exception condition is if the subject contains "fun" or "chat" 
 
 Set oExceptSubject = _ 
 
 oRule.Exceptions.Subject 
 
 With oExceptSubject 
 
 .Enabled = True 
 
 .Text = Array("fun", "chat") 
 
 End With 
 
 
 
 'Update the server and display progress dialog 
 
 colRules.Save 
 
End Sub
```


## See also


#### Concepts


[Rules Object](rules-object-outlook.md)

