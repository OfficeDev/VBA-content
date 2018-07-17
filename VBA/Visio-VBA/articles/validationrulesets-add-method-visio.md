---
title: ValidationRuleSets.Add Method (Visio)
keywords: vis_sdr.chm18116005
f1_keywords:
- vis_sdr.chm18116005
ms.prod: visio
api_name:
- Visio.ValidationRuleSets.Add
ms.assetid: 69e2526a-e787-d9a8-45c1-e2f1e48faa03
ms.date: 06/08/2017
---


# ValidationRuleSets.Add Method (Visio)

Adds a new, empty  **[ValidationRuleSet](validationruleset-object-visio.md)** object to the **[ValidationRuleSets](validationrulesets-object-visio.md)** collection of the document.


## Syntax

 _expression_ . **Add**( **_NameU_** )

 _expression_ A variable that represents a **ValidationRuleSet** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _NameU_|Required| **String**|The universal name to assign to the new validation rule set.|

### Return Value

 **ValidationRuleSet**


## Remarks

If the  _NameU_ parameter is not a valid string or if it matches the universal name of an existing rule set in the document, Microsoft Visio returns an Invalid Parameter error.

The default property values of the new validation rule set are as follows:  **[Description](validationruleset-description-property-visio.md)** = [empty]; **Enabled** = **True** ; **[Name](validationruleset-name-property-visio.md)** = **NameU** ; **[RuleSetFlags](validationruleset-rulesetflags-property-visio.md)** = **visRuleSetDefault** .


## Example

The following Visual Basic for Applications (VBA) example shows how to use the  **Add** method to add a validation rule set named "Connectivity" to the active document.


```vb
Set vsoDocument = Visio.ActiveDocument

' Add a validation rule set to the document.
Set vsoValidationRuleSet = 
vsoDocument.Validation.RuleSets.Add("Connectivity")

```


