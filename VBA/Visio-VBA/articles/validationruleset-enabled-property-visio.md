---
title: ValidationRuleSet.Enabled Property (Visio)
keywords: vis_sdr.chm18213455
f1_keywords:
- vis_sdr.chm18213455
ms.prod: visio
api_name:
- Visio.ValidationRuleSet.Enabled
ms.assetid: 1fc9c692-736d-6686-fb47-5bd7efb39773
ms.date: 06/08/2017
---


# ValidationRuleSet.Enabled Property (Visio)

Determines whether the rules in the specified validation rule set are checked when validation is triggered for the current document. Read/write.


## Syntax

 _expression_ . **Enabled**

 _expression_ A variable that represents a **[ValidationRuleSet](validationruleset-object-visio.md)** object.


### Return Value

 **Boolean**


## Remarks

If the value of the  **Enabled** property is **True** , the rules in the validation rule set are checked when validation is triggered for the current document. Validation is triggered when the user clicks **Check Diagram** on the **Process** tab or when the **[Validate](validation-validate-method-visio.md)** method is run on the current document.

Rule sets for which the value of  **Enabled** is **False** are purged from the current document when the **[RemoveHiddenInformation](document-removehiddeninformation-method-visio.md)** method is run with the **visRHIValidationRules** flag set, or when the equivalent command is issued in the user interface.


## Example

The following Visual Basic for Applications (VBA) example shows how to use the  **Enabled** property to enable a validation rule set named "Connectivity" in the active document.


```vb
Set vsoDocument = Visio.ActiveDocument

Set vsoValidationRuleSet = vsoDocument.Validation.RuleSets.Add("Connectivity")

' Enable the rule set.
vsoValidationRuleSet.Enabled = True
```


