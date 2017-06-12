---
title: Rules Object (Outlook)
keywords: vbaol11.chm3160
f1_keywords:
- vbaol11.chm3160
ms.prod: outlook
api_name:
- Outlook.Rules
ms.assetid: dd41b4de-bf5f-5532-46c9-394a5d078bec
ms.date: 06/08/2017
---


# Rules Object (Outlook)

Represents a set of  **[Rule](rule-object-outlook.md)** objects that are the rules available in the current session.


## Remarks

The Rules object model consists primarily of these objects:  **Rules**, **Rule**, **[RuleActions](http://msdn.microsoft.com/library/82ba76cd-86a4-3372-cb51-2df1d58c8b71%28Office.15%29.aspx)**, **[RuleConditions](ruleconditions-object-outlook.md)**, **[RuleAction](http://msdn.microsoft.com/library/6451788f-e5ed-239c-a34d-b564b52d8955%28Office.15%29.aspx)**, **[RuleCondition](http://msdn.microsoft.com/library/e03f91c2-2c08-b036-104a-d6246f28bc2d%28Office.15%29.aspx)**, and the derived objects for certain rule actions and rule conditions. It provides partial parity with the Rules and Alerts Wizard in the Outlook user interface. Although it does not support creation of every single rule that you can possibly create using the Wizard, it supports the most commonly used rule actions and conditions.

For more information on how to programmatically create, edit, and delete rules, see [Managing Rules in the Outlook Object Model](http://msdn.microsoft.com/library/05ddd643-e9bd-a37d-b680-b8519960a5f6%28Office.15%29.aspx) and[How to: Create a Rule to Move Specific E-mails to a Folder](http://msdn.microsoft.com/library/e72fa307-8224-c2d2-1318-a18cd8e9f22f%28Office.15%29.aspx).


## Methods



|**Name**|
|:-----|
|[Create](http://msdn.microsoft.com/library/84789ccc-a6c2-9f79-5338-45b03b116dd5%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/fe696181-9f61-0eb7-9634-5f7c007f1606%28Office.15%29.aspx)|
|[Remove](http://msdn.microsoft.com/library/6d4bb971-b38a-0434-1b6a-8892689549d6%28Office.15%29.aspx)|
|[Save](http://msdn.microsoft.com/library/d838eca0-4ec5-ab43-a031-fd65ab7d9f3c%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/406b1f7c-1714-5f0e-5d9f-37ddc963ca69%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/3ee88b9e-4cb3-c80b-6386-4b35ef59d27b%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/b1753709-5693-9f2a-cd11-0e3c4e5e0982%28Office.15%29.aspx)|
|[IsRssRulesProcessingEnabled](http://msdn.microsoft.com/library/7eff75e6-1e1a-0fbf-9d05-2f40e7f08145%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/4a74aeb3-4502-a59f-fdb9-29d7181f3bb3%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/c544e009-623c-3e4d-b71a-9177dcfcc668%28Office.15%29.aspx)|

## See also


#### Other resources


[Rules Object Members](http://msdn.microsoft.com/library/39fb5418-ff5a-1714-d3b5-07cc28893821%28Office.15%29.aspx)
[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
