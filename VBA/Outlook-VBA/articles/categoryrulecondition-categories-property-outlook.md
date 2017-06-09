---
title: CategoryRuleCondition.Categories Property (Outlook)
keywords: vbaol11.chm2446
f1_keywords:
- vbaol11.chm2446
ms.prod: outlook
api_name:
- Outlook.CategoryRuleCondition.Categories
ms.assetid: 7662a095-43e9-7668-f6f7-d0701b87b28c
ms.date: 06/08/2017
---


# CategoryRuleCondition.Categories Property (Outlook)

Returns or sets an array of  **String** elements that represents the categories evaluated by the rule condition. Read/write.


## Syntax

 _expression_ . **Categories**

 _expression_ A variable that represents a **CategoryRuleCondition** object.


## Remarks

You can assign an array with one element to evaluate a single category or an array of multiple strings to evaluate multiple categories. Multiple category strings assigned in an array are evaluated using the logical OR operation.

This property uses the character specified in the value name,  **sList** , under **HKEY_CURRENT_USER\Control Panel\International** in the Windows registry, as the delimiter for multiple categories.

If a string specified by  **Categories** matches a category of the message, the condition evaluates to **True** .

Outlook does not check to determine if the  **Categories** property contains category names that are in the master category list.

Returns an error if  **Categories** contains one or more empty strings.


## See also


#### Concepts


[CategoryRuleCondition Object](categoryrulecondition-object-outlook.md)

