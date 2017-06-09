---
title: RuleConditions Object (Outlook)
keywords: vbaol11.chm3172
f1_keywords:
- vbaol11.chm3172
ms.prod: outlook
api_name:
- Outlook.RuleConditions
ms.assetid: e8e9a05a-b36b-add2-b294-8cdc5a97e119
ms.date: 06/08/2017
---


# RuleConditions Object (Outlook)

Contains a set of  **[RuleCondition](http://msdn.microsoft.com/library/e03f91c2-2c08-b036-104a-d6246f28bc2d%28Office.15%29.aspx)** objects or objects derived from **RuleCondition**, representing the conditions or exception conditions that must be satisfied in order for the **[Rule](rule-object-outlook.md)** to execute.


## Remarks

The  **RuleConditions** object include both rule conditions and rule exceptions. The type of rule condition that can be added to a **RuleConditions** collection depends upon the **[Rule.RuleType](http://msdn.microsoft.com/library/6ae3ca3c-860e-9cbd-d0d0-c36039b54c39%28Office.15%29.aspx)**.

The  **RuleConditions** object is a fixed collection. A **RuleCondition** object or a type that is derived from the **RuleCondition** object cannot be added or removed from the **RuleConditions** object.

The Rules object model provides partial parity with the Rules and Alerts Wizard in the Outlook user interface. It supports the most commonly used rule actions and conditions. Although it does not support creating rules with any rule action or rule condition that the Wizard supports, you can still enumerate and enable these rule actions and conditions in existing rules. 

For more information on rule conditions, see [Specifying Rule Conditions](http://msdn.microsoft.com/library/812c131a-fe23-1b8b-5e2d-9459d7102630%28Office.15%29.aspx) and[How to: Create a Rule to Move Specific E-mails to a Folder](http://msdn.microsoft.com/library/e72fa307-8224-c2d2-1318-a18cd8e9f22f%28Office.15%29.aspx).


## Methods



|**Name**|
|:-----|
|[Item](http://msdn.microsoft.com/library/2fc986a5-e77a-e8c9-b8bf-4af85720a771%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Account](http://msdn.microsoft.com/library/9e1ecf7d-b832-e657-92df-42bb28f5d924%28Office.15%29.aspx)|
|[AnyCategory](http://msdn.microsoft.com/library/b174ad44-570b-fa6f-1abc-452929dd2154%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/c8e620fa-eff1-4e21-e547-a3bc690cf853%28Office.15%29.aspx)|
|[Body](http://msdn.microsoft.com/library/b962167e-b1d6-045c-79b7-0ba4c96b123c%28Office.15%29.aspx)|
|[BodyOrSubject](http://msdn.microsoft.com/library/ced8a26a-9a54-d1f4-18f6-dd52a8203892%28Office.15%29.aspx)|
|[Category](http://msdn.microsoft.com/library/f1131bf8-4752-4e93-c68d-73c0511d22da%28Office.15%29.aspx)|
|[CC](http://msdn.microsoft.com/library/0475c994-4887-f268-d7f7-46b3c4e7186c%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/d4072c77-2906-e26c-5d1a-a88969a95262%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/7950c105-4528-40aa-f263-b800a68ae1ad%28Office.15%29.aspx)|
|[FormName](http://msdn.microsoft.com/library/9f292443-1af7-500e-2959-1fce4c7d4824%28Office.15%29.aspx)|
|[From](http://msdn.microsoft.com/library/3ebda0d0-ba44-95c6-ed02-a9c6acbf1f1c%28Office.15%29.aspx)|
|[FromAnyRSSFeed](http://msdn.microsoft.com/library/df580ca7-ee2f-9c3a-ebc7-ca35528554cd%28Office.15%29.aspx)|
|[FromRssFeed](http://msdn.microsoft.com/library/ef312495-4d65-bb89-c543-59c5473171ff%28Office.15%29.aspx)|
|[HasAttachment](http://msdn.microsoft.com/library/d480c5ff-2313-f428-88b6-0cf52ffb4003%28Office.15%29.aspx)|
|[Importance](http://msdn.microsoft.com/library/619fc6e3-7a4e-dc00-9108-857d383f460e%28Office.15%29.aspx)|
|[MeetingInviteOrUpdate](http://msdn.microsoft.com/library/0204dfdb-bf93-db11-3550-3b23fdec47c9%28Office.15%29.aspx)|
|[MessageHeader](http://msdn.microsoft.com/library/311f8834-f12b-50db-1f0d-00d6ebed7e9d%28Office.15%29.aspx)|
|[NotTo](http://msdn.microsoft.com/library/9889e503-05cd-ebf8-40e0-358327798b6a%28Office.15%29.aspx)|
|[OnLocalMachine](http://msdn.microsoft.com/library/747de02c-d76d-9da3-c582-50719e618eb4%28Office.15%29.aspx)|
|[OnlyToMe](http://msdn.microsoft.com/library/208e7bc4-2938-ecc8-7af5-9e3e256fe5b1%28Office.15%29.aspx)|
|[OnOtherMachine](http://msdn.microsoft.com/library/03d96697-5978-8a0c-7356-dfe721f5b05d%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/0f0b6104-0bb1-404c-eae3-6881d80dc690%28Office.15%29.aspx)|
|[RecipientAddress](http://msdn.microsoft.com/library/1b8f361e-0481-75dc-e66e-2bc69228773a%28Office.15%29.aspx)|
|[SenderAddress](http://msdn.microsoft.com/library/6e5eb1cc-385f-b1b2-aea7-12629cc31030%28Office.15%29.aspx)|
|[SenderInAddressList](http://msdn.microsoft.com/library/bf836af6-fd72-d77d-dfbe-90a8038188a6%28Office.15%29.aspx)|
|[SentTo](http://msdn.microsoft.com/library/54039c2f-b2a5-2878-84c0-b129b4ce96fa%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/0a214009-1bd1-9631-a80c-e942680ae878%28Office.15%29.aspx)|
|[Subject](http://msdn.microsoft.com/library/d6d51efb-9eec-0c07-ca8f-616791822f91%28Office.15%29.aspx)|
|[ToMe](http://msdn.microsoft.com/library/c1b4a68a-64da-c0e8-00a7-11f49f995934%28Office.15%29.aspx)|
|[ToOrCc](http://msdn.microsoft.com/library/28a34223-e47d-3843-4648-fe757568d406%28Office.15%29.aspx)|

## See also


#### Other resources


[RuleConditions Object Members](http://msdn.microsoft.com/library/b2af6ebf-f9f8-8106-20a3-1725c3b78174%28Office.15%29.aspx)
[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
