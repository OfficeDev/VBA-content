---
title: ValidationRule.TargetType Property (Visio)
keywords: vis_sdr.chm18462660
f1_keywords:
- vis_sdr.chm18462660
ms.prod: visio
api_name:
- Visio.ValidationRule.TargetType
ms.assetid: 818e47b6-7832-e9a3-9e29-34bd50d466b4
ms.date: 06/08/2017
---


# ValidationRule.TargetType Property (Visio)

Determines the type of object to which the validation rule applies. Read/write.


## Syntax

 _expression_ . **TargetType**

 _expression_ A variable that represents a **[ValidationRule](validationrule-object-visio.md)** object.


### Return Value

 **[VisRuleTargets](visruletargets-enumeration-visio.md)**


## Remarks

Valid validation-rule targets include documents, pages, and shapes. The  **TargetType** property value must be one of the following **VisRuleTargets** constants.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visRuleTargetShape**|0|The rule applies to shapes in the document.|
| **visRuleTargetPage**|1|The rule applies to pages in the document.|
| **visRuleTargetDocument**|2|The rule applies to the document itself.|
If you pass any other value to the  **TargetType** property, Visio returns an invalid-parameter error.


## Example

The following Visual Basic for Applications (VBA) example shows how to use the  **TargetType** property to specify the type of object to which the validation rule named "Unglued2DShape" should apply.


```vb
Set vsoValidationRule = vsoValidationRuleSet.Rules.Add("Unglued2DShape")
vsoValidationRule.TargetType = Visio.VisRuleTargets.visRuleTargetShape
```


