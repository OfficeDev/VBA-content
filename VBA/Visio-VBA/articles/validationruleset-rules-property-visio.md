---
title: ValidationRuleSet.Rules Property (Visio)
keywords: vis_sdr.chm18262645
f1_keywords:
- vis_sdr.chm18262645
ms.prod: visio
api_name:
- Visio.ValidationRuleSet.Rules
ms.assetid: 7890ca86-74b3-1dd6-8322-f3fbde235115
ms.date: 06/08/2017
---


# ValidationRuleSet.Rules Property (Visio)

Returns the collection of validation rules in the validation rule set. Read-only.


## Syntax

 _expression_ . **Rules**

 _expression_ A variable that represents a **[ValidationRuleSet](validationruleset-object-visio.md)** object.


### Return Value

 **[ValidationRules](validationrules-object-visio.md)**


## Example

The following sample code is based on code provided by: [David Parker](http://www.bvisual.net)

The following Visual Basic for Applications (VBA) example shows how to use the  **Rules** property to get the names of all the validation rules in an existing validation rule set named "Fault Tree Analysis" in the active document. The example then prints those names in the **Immediate** window.




```vb
Public Sub Rules_Example()

    Dim vsoDocument As Visio.Document
    Dim vsoValidationRule As Visio.ValidationRule
    Dim strValidationRuleSetNameU As String
    
    Set vsoDocument = Visio.ActiveDocument
    strValidationRuleSetNameU = "FaultTreeAnalysis"
    
    For Each vsoValidationRule In vsoDocument.Validation.RuleSets(strValidationRuleSetNameU).Rules
       
        Debug.Print vsoValidationRule.NameU
       
    Next

End Sub
```


