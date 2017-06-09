---
title: ValidationRule.FilterExpression Property (Visio)
keywords: vis_sdr.chm18462655
f1_keywords:
- vis_sdr.chm18462655
ms.prod: visio
api_name:
- Visio.ValidationRule.FilterExpression
ms.assetid: bbca9cf8-ad34-062b-eaf5-b30a943db1b1
ms.date: 06/08/2017
---


# ValidationRule.FilterExpression Property (Visio)

Gets or sets the logical expression that determines whether the validation rule should be applied to a target object. Read/write.


## Syntax

 _expression_ . **FilterExpression**

 _expression_ A variable that represents a **[ValidationRule](validationrule-object-visio.md)** object.


### Return Value

 **String**


## Remarks

When you validate a diagram by calling the  **[Validate](validation-validate-method-visio.md)** method or by clicking **Check Diagram** on the **Process** tab, Microsoft Visio uses the expression that you set as the **FilterExpression** property value to determine whether a target object must satisfy the validation rule. If the filter expression you set evaluates to **True** , Visio uses the **[TestExpression](validationrule-testexpression-property-visio.md)** property value you set to determine whether to generate an issue for the target object. If the filter expression evaluates to **False** , Visio does not apply the validation rule to the target object during validation.

The syntax for the  **FilterExpression** property value is the same as that for a ShapeSheet expression. When you set the **FilterExpression** property, Visio does not validate the syntax of the filter expression. If the expression is not syntactically correct, Visio does not apply the validation rule to the target object during validation


## Example

The following Visual Basic for Applications (VBA) example shows how to use the  **FilterExpression** property to determine whether a particular shape target must satisfy a validation rule.


```vb
' The validation function Is1D() returns a Boolean value that 
' indicates whether the shape is 1D (True) or 2D (False).
vsoValidationRule.FilterExpression = "NOT(Is1D())"
```


