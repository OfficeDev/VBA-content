---
title: Rule.IsLocalRule Property (Outlook)
keywords: vbaol11.chm2172
f1_keywords:
- vbaol11.chm2172
ms.prod: outlook
api_name:
- Outlook.Rule.IsLocalRule
ms.assetid: 430a8240-8572-5b9a-5e59-2b38bb1b3d17
ms.date: 06/08/2017
---


# Rule.IsLocalRule Property (Outlook)

Returns a  **Boolean** that indicates if the rule executes as a client-side rule. Read-only.


## Syntax

 _expression_ . **IsLocalRule**

 _expression_ A variable that represents a **Rule** object.


## Remarks

A client-side rule executes only when Outlook is running. If  **IsLocalRule** is **False** , then the rule executes as a server-side rule.

If you have a Microsoft Exchange Server e-mail account, the server can apply server-side rules to your messages even if you don't have Outlook running. The rules must be set to be applied to messages when they are delivered to your Inbox on the server, and the rules must be able to run to completion on the server. For example, a rule cannot run to completion on the server if the action specifies that a message be printed. If a rule cannot be applied on the server, it is applied when you start Outlook.

If the rules collection on a store contains both server and client-side rules, the server-side rules are applied first, followed by the client-side rules.


## See also


#### Concepts


[Rule Object](rule-object-outlook.md)

