---
title: RuleActions Object (Outlook)
keywords: vbaol11.chm3162
f1_keywords:
- vbaol11.chm3162
ms.prod: outlook
api_name:
- Outlook.RuleActions
ms.assetid: 82ba76cd-86a4-3372-cb51-2df1d58c8b71
ms.date: 06/08/2017
---


# RuleActions Object (Outlook)

The  **RuleActions** object contains a set of **[RuleAction](ruleaction-object-outlook.md)** objects or objects derived from **RuleAction**, representing the actions that are executed on a **[Rule](rule-object-outlook.md)** object.


## Remarks

The  **RuleActions** object is a fixed collection. **RuleAction** objects or types that derive from the **RuleAction** object cannot be added to or removed from the **RuleActions** object.

The Rules object model provides partial parity with the Rules and Alerts Wizard in the Outlook user interface. It supports the most commonly used rule actions and conditions. Although it does not support creating rules with any rule action or rule condition that the Wizard supports, you can still enumerate and enable these rule actions and conditions in existing rules. 

For more information on rule actions, see [Specifying Rule Actions](http://msdn.microsoft.com/library/c5f83c81-0e01-38aa-5ec7-3932b4443e43%28Office.15%29.aspx) and[How to: Create a Rule to Move Specific E-mails to a Folder](http://msdn.microsoft.com/library/e72fa307-8224-c2d2-1318-a18cd8e9f22f%28Office.15%29.aspx).


## Methods



|**Name**|
|:-----|
|[Item](ruleactions-item-method-outlook.md)|

## Properties



|**Name**|
|:-----|
|[Application](ruleactions-application-property-outlook.md)|
|[AssignToCategory](ruleactions-assigntocategory-property-outlook.md)|
|[CC](ruleactions-cc-property-outlook.md)|
|[Class](ruleactions-class-property-outlook.md)|
|[ClearCategories](ruleactions-clearcategories-property-outlook.md)|
|[CopyToFolder](ruleactions-copytofolder-property-outlook.md)|
|[Count](ruleactions-count-property-outlook.md)|
|[Delete](ruleactions-delete-property-outlook.md)|
|[DeletePermanently](ruleactions-deletepermanently-property-outlook.md)|
|[DesktopAlert](ruleactions-desktopalert-property-outlook.md)|
|[Forward](ruleactions-forward-property-outlook.md)|
|[ForwardAsAttachment](ruleactions-forwardasattachment-property-outlook.md)|
|[MarkAsTask](ruleactions-markastask-property-outlook.md)|
|[MoveToFolder](ruleactions-movetofolder-property-outlook.md)|
|[NewItemAlert](ruleactions-newitemalert-property-outlook.md)|
|[NotifyDelivery](ruleactions-notifydelivery-property-outlook.md)|
|[NotifyRead](ruleactions-notifyread-property-outlook.md)|
|[Parent](ruleactions-parent-property-outlook.md)|
|[PlaySound](ruleactions-playsound-property-outlook.md)|
|[Redirect](ruleactions-redirect-property-outlook.md)|
|[Session](ruleactions-session-property-outlook.md)|
|[Stop](ruleactions-stop-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
