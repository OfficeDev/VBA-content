---
title: AddressRuleCondition.Address Property (Outlook)
keywords: vbaol11.chm2955
f1_keywords:
- vbaol11.chm2955
ms.prod: outlook
api_name:
- Outlook.AddressRuleCondition.Address
ms.assetid: de4186ec-0741-8ff6-7789-af0a46c470e0
ms.date: 06/08/2017
---


# AddressRuleCondition.Address Property (Outlook)

Returns or sets an array of  **String** elements to evaluate the address rule condition. Read/write.


## Syntax

 _expression_ . **Address**

 _expression_ A variable that represents an **AddressRuleCondition** object.


## Remarks

You can assign an array with one element to evaluate a single address or an array of multiple strings to evaluate multiple addresses. Multiple address strings assigned in an array are evaluated using the logical OR operation.

If a string specified by  **Address** is contained in the recipient or sender address, the condition evaluates to **True** .


## See also


#### Concepts


[AddressRuleCondition Object](addressrulecondition-object-outlook.md)

