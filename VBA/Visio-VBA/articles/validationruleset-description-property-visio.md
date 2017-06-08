---
title: ValidationRuleSet.Description Property (Visio)
keywords: vis_sdr.chm18213405
f1_keywords:
- vis_sdr.chm18213405
ms.prod: visio
api_name:
- Visio.ValidationRuleSet.Description
ms.assetid: 65083a0d-66bf-0395-6ecb-db8de13a766e
ms.date: 06/08/2017
---


# ValidationRuleSet.Description Property (Visio)

Specifies the description of the  **[ValidationRuleSet](validationruleset-object-visio.md)** object that appears in the user interface. Read/write.


## Syntax

 _expression_ . **Description**

 _expression_ A variable that represents a **ValidationRuleSet** object.


### Return Value

 **String**


## Remarks

You cannot set the  **Description** property to a value that exceeds 255 characters.


## Example

The following Visual Basic for Applications (VBA) example shows how to use the  **Description** property to set the description of a validation rule set named "Connectivity" in the active document.


```vb
Set vsoDocument = Visio.ActiveDocument
Set vsoValidationRuleSet = vsoDocument.Validation.RuleSets.Add("Connectivity")
vsoValidationRuleSet.Description = "Verify that shapes are correctly connected in the document."
```


