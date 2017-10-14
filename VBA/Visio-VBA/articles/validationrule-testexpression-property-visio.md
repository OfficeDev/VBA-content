---
title: ValidationRule.TestExpression Property (Visio)
keywords: vis_sdr.chm18462665
f1_keywords:
- vis_sdr.chm18462665
ms.prod: visio
api_name:
- Visio.ValidationRule.TestExpression
ms.assetid: 0d780351-ca46-e896-c6a4-5ae899427ae0
ms.date: 06/08/2017
---


# ValidationRule.TestExpression Property (Visio)

Gets or sets the logical expression that determines whether the target object satisfies the validation rule. Read/write.


## Syntax

 _expression_ . **TestExpression**

 _expression_ A variable that represents a **[ValidationRule](validationrule-object-visio.md)** object.


### Return Value

 **String**


## Remarks

When you validate a diagram by calling the  **[Validate](validation-validate-method-visio.md)** method or by clicking **Check Diagram** on the **Process** tab, Microsoft Visio uses the test expression that you set as the **TestExpression** property value to determine whether the target object satisfies the validation rule. If the test expression evaluates to **False** , Visio generates a validation issue. If the test expression evaluates to **True** , no validation issue is generated.

Visio evaluates the test expression for target objects only when the value of the  **[FilterExpression](validationrule-filterexpression-property-visio.md)** property of the **ValidationRule** object evaluates to **True** .

The syntax for the  **TestExpression** property value is the same as that for a ShapeSheet expression. When you set the **TestExpression** property value, Visio does not check the syntax of the test expression. If the test expression is not syntactically correct, the evaluation of the expression fails during validation and Visio generates a validation issue at that time.


## Example

The following Visual Basic for Applications (VBA) example shows how to use the  **TestExpression** property to determine whether a particular shape target satisfies a validation rule.


```vb
' Add a validation rule to the document.
Set vsoValidationRule = vsoValidationRuleSet.Rules.Add("Unglued2DShape")
vsoValidationRule.Category = "Shapes"
vsoValidationRule.Description = "This 2-dimensional shape is not connected to
any other shape."
vsoValidationRule.Ignored = False
vsoValidationRule.TargetType = Visio.VisRuleTargets.visRuleTargetShape

' The validation function Is1D() returns a Boolean value that indicates 
' whether the shape is 1D (True) or 2D (False).
vsoValidationRule.FilterExpression = "NOT(Is1D())"

' The validation function GLUEDSHAPES returns a set of 
' shapes glued to the shape.
' It takes as input one parameter that indicates the direction of the glue.
' The direction values are equivalent to members of VisGluedShapesFlags:
' 0 = visGluedShapesAll1D, and 3 = visGluedShapesAll2D
' It takes as input one parameter indicating the direction of the glue.

' The validation function AGGCOUNT takes a set of shapes as its input, and 
' returns the number of shapes in the set.
vsoValidationRule.TestExpression = "AGGCOUNT(GLUEDSHAPES(0)) + AGGCOUNT(GLUEDSHAPES(3)) > 0"
```


