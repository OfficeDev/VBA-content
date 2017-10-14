---
title: ValidationRule.Description Property (Visio)
keywords: vis_sdr.chm18413405
f1_keywords:
- vis_sdr.chm18413405
ms.prod: visio
api_name:
- Visio.ValidationRule.Description
ms.assetid: 111e41fd-f6ea-c33e-a4f3-18d609e16ad1
ms.date: 06/08/2017
---


# ValidationRule.Description Property (Visio)

Specifies the description of the  **[ValidationRule](validationrule-object-visio.md)** object that appears in the user interface. Read/write.


## Syntax

 _expression_ . **Description**

 _expression_ A variable that represents a **ValidationRule** object.


### Return Value

 **String**


## Remarks

You cannot set the  **Description** property to a value that exceeds 255 characters.


## Example

The following Visual Basic for Applications (VBA) example shows how to use the  **Description** property to set the description that appears in the user interface for the validation rule named "Unglued2DShape".


```vb
Set vsoValidationRule = vsoValidationRuleSet.Rules.Add("Unglued2DShape")
vsoValidationRule.Description = "This 2-dimensional shape is not connected to any other shape."
```


