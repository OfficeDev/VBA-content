---
title: FormNameRuleCondition.FormName Property (Outlook)
keywords: vbaol11.chm2454
f1_keywords:
- vbaol11.chm2454
ms.prod: outlook
api_name:
- Outlook.FormNameRuleCondition.FormName
ms.assetid: 993f2ee0-58eb-bed0-5819-11148792b8f0
ms.date: 06/08/2017
---


# FormNameRuleCondition.FormName Property (Outlook)

Returns or sets an  **Object** that represents an array of form identifiers to be evaluated by the rule condition. Read/write.


## Syntax

 _expression_ . **FormName**

 _expression_ A variable that represents a **FormNameRuleCondition** object.


## Remarks

Even though the Rules and Alerts Wizard uses the display name of a form as an identifier, programmatically,  **FormName** uses the message class of the form as an identifier.

You can assign an array with one string or an array of multiple strings to the  **FormName** property. Multiple form identifiers assigned in an array are evaluated using the logical OR operation.

 **FormName** returns an error if it contains one or more empty strings.


## See also


#### Concepts


[FormNameRuleCondition Object](formnamerulecondition-object-outlook.md)

