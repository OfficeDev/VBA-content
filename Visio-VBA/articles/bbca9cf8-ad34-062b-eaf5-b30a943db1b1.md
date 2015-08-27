
# ValidationRule.FilterExpression Property (Visio)

 **Last modified:** July 28, 2015

 _**Applies to:** Visio 2013 Preview_

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Gets or sets the logical expression that determines whether the validation rule should be applied to a target object. Read/write.

## Syntax
<a name="sectionSection0"> </a>

 _expression_. **FilterExpression**

 _expression_A variable that represents a  ** [ValidationRule](c9efb9b4-10b0-b6aa-cc78-2a01fd3e8357.md)** object.


### Return Value

 **String**


## Remarks
<a name="sectionSection1"> </a>

When you validate a diagram by calling the  ** [Validate](9e8b8bcd-674e-c7ac-543c-027ed02519cd.md)** method or by clicking **Check Diagram** on the **Process** tab, Microsoft Visio uses the expression that you set as the **FilterExpression** property value to determine whether a target object must satisfy the validation rule. If the filter expression you set evaluates to **True**, Visio uses the  ** [TestExpression](0d780351-ca46-e896-c6a4-5ae899427ae0.md)** property value you set to determine whether to generate an issue for the target object. If the filter expression evaluates to **False**, Visio does not apply the validation rule to the target object during validation.

The syntax for the  **FilterExpression** property value is the same as that for a ShapeSheet expression. When you set the **FilterExpression** property, Visio does not validate the syntax of the filter expression. If the expression is not syntactically correct, Visio does not apply the validation rule to the target object during validation


## Example
<a name="sectionSection2"> </a>

The following Visual Basic for Applications (VBA) example shows how to use the  **FilterExpression** property to determine whether a particular shape target must satisfy a validation rule.


```
' The validation function Is1D() returns a Boolean value that 
' indicates whether the shape is 1D (True) or 2D (False).
vsoValidationRule.FilterExpression = "NOT(Is1D())"
```

