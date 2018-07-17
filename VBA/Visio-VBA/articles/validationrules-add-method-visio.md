---
title: ValidationRules.Add Method (Visio)
keywords: vis_sdr.chm18316005
f1_keywords:
- vis_sdr.chm18316005
ms.prod: visio
api_name:
- Visio.ValidationRules.Add
ms.assetid: 14b0ab24-5ff6-cde5-8311-ccf2989712c9
ms.date: 06/08/2017
---


# ValidationRules.Add Method (Visio)

Adds a new, empty  **[ValidationRule](validationrule-object-visio.md)** object to the **[ValidationRules](validationrules-object-visio.md)** collection of the document.


## Syntax

 _expression_ . **Add**( **_NameU_** )

 _expression_ A variable that represents a **ValidationRules** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _NameU_|Required| **String**|The universal name to assign to the new validation rule.|

### Return Value

 **ValidationRule**


## Remarks

If the  _NameU_ parameter is not a valid string, Visio returns an Invalid Parameter error.

The default property values of the new validation rule are as follows:  **[Category](validationrule-category-property-visio.md)** = [empty]; **[Description](validationrule-description-property-visio.md)** = "Unknown"; **[FilterExpression](validationrule-filterexpression-property-visio.md)** = [empty]; **[Ignored](validationrule-ignored-property-visio.md)** = **False** ; **[TargetType](validationrule-targettype-property-visio.md)** = **visRuleTargetShape** ; **[TestExpression](validationrule-testexpression-property-visio.md)** = [empty].


## Example

The following sample code is based on code provided by: [David Parker](http://www.bvisual.net)

The following Visual Basic for Applications (VBA) example shows how to use the  **Add** method to add a new validation rule named "UngluedConnector" to an existing validation rule set named "Fault Tree Analysis" in the active document.




```vb
Public Sub Add_Example()

    Dim vsoValidationRule As Visio.ValidationRule
    Dim vsoValidationRuleSet As Visio.ValidationRuleSet
    Dim strValidationRuleSetNameU As String
    Dim strValidationRuleNameU As String
    
    strValidationRuleSetNameU = "Fault Tree Analysis"
    strValidationRuleNameU = "UngluedConnector"
    
    Set vsoValidationRuleSet = ActiveDocument.Validation.RuleSets(strValidationRuleSetNameU)
    Set vsoValidationRule = vsoValidationRuleSet.Rules.Add(strValidationRuleNameU)

End Sub
```


