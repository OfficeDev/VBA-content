---
title: Rule Object (Outlook)
keywords: vbaol11.chm3161
f1_keywords:
- vbaol11.chm3161
ms.prod: outlook
api_name:
- Outlook.Rule
ms.assetid: ea2ddbcc-fd65-a636-c6da-79950033f385
ms.date: 06/08/2017
---


# Rule Object (Outlook)

Represents an Outlook rule.


## Remarks

Both client and server side rules are represented by the  **Rule** object.

The Rules object model consists primarily of these objects:  **[Rules](rules-object-outlook.md)**, **Rule**, **[RuleActions](http://msdn.microsoft.com/library/82ba76cd-86a4-3372-cb51-2df1d58c8b71%28Office.15%29.aspx)**, **[RuleConditions](ruleconditions-object-outlook.md)**, **[RuleAction](http://msdn.microsoft.com/library/6451788f-e5ed-239c-a34d-b564b52d8955%28Office.15%29.aspx)**, **[RuleCondition](http://msdn.microsoft.com/library/e03f91c2-2c08-b036-104a-d6246f28bc2d%28Office.15%29.aspx)**, and the derived objects for certain rule actions and rule conditions. It provides partial parity with the Rules and Alerts Wizard in the Outlook user interface. Although it does not support creation of every single rule that you can possibly create using the Wizard, it supports the most commonly used rule actions and conditions.

For more information on how to programmatically create, edit, and delete rules, see [Manage Rules in the Outlook Object Model](http://msdn.microsoft.com/library/05ddd643-e9bd-a37d-b680-b8519960a5f6%28Office.15%29.aspx) and[How to: Create a Rule to Move Specific E-mails to a Folder](http://msdn.microsoft.com/library/e72fa307-8224-c2d2-1318-a18cd8e9f22f%28Office.15%29.aspx).


## Methods



|**Name**|
|:-----|
|[Execute](http://msdn.microsoft.com/library/487abb6f-9003-04a4-f4e2-3f66b3ba5a52%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Actions](http://msdn.microsoft.com/library/2b1e2ad4-c735-b3a8-6b27-5004f10393ce%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/8c21ce34-b206-315c-16ff-e27bfc606d85%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/9d32cc3e-f17f-aaa8-f08c-ccef85f387ce%28Office.15%29.aspx)|
|[Conditions](http://msdn.microsoft.com/library/e2cacf1c-95eb-31d3-012c-7cf9426053d5%28Office.15%29.aspx)|
|[Enabled](http://msdn.microsoft.com/library/9ba65f87-799f-7a22-04a1-c0abcb320559%28Office.15%29.aspx)|
|[Exceptions](http://msdn.microsoft.com/library/843c2690-ee39-bac7-d593-80c3dd31087f%28Office.15%29.aspx)|
|[ExecutionOrder](http://msdn.microsoft.com/library/070d50ca-4b0b-5629-1609-81ab8a3620d1%28Office.15%29.aspx)|
|[IsLocalRule](http://msdn.microsoft.com/library/430a8240-8572-5b9a-5e59-2b38bb1b3d17%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/6c559ffe-b25c-ff49-31d1-1fd44935a8f3%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/d8b810ee-76c6-9aa4-68ca-97a62a35c81c%28Office.15%29.aspx)|
|[RuleType](http://msdn.microsoft.com/library/6ae3ca3c-860e-9cbd-d0d0-c36039b54c39%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/7502f919-cf8f-d795-87b1-9812c0d150d1%28Office.15%29.aspx)|

## See also


#### Other resources


[Rule Object Members](http://msdn.microsoft.com/library/29a5f487-dbcc-7312-c8ba-a05199ce8513%28Office.15%29.aspx)
[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
