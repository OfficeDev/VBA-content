---
title: ValidationRule.Delete Method (Visio)
keywords: vis_sdr.chm18416165
f1_keywords:
- vis_sdr.chm18416165
ms.prod: visio
api_name:
- Visio.ValidationRule.Delete
ms.assetid: 3727bddb-57b9-23c3-d12d-f47d43260087
ms.date: 06/08/2017
---


# ValidationRule.Delete Method (Visio)

Deletes the  **[ValidationRule](validationrule-object-visio.md)** object from the document.


## Syntax

 _expression_ . **Delete**

 _expression_ A variable that represents a **ValidationRule** object.


### Return Value

 **Nothing**


## Remarks

Calling the  **Delete** method also deletes all **[ValidationIssue](validationissue-object-visio.md)** objects that are associated with the validation rule.


## Example

The following sample code is based on code provided by: [David Parker](http://www.bvisual.net)

The following Visual Basic for Applications (VBA) example shows how to use the  **Delete** method to delete a validation rule named "Unglued Connector" from the validation rule set named "Fault Tree Analysis" in the active document.




```vb
' Delete a named rule from a named rule set.
Public Sub Delete_Example()

    Dim strValidationRuleSetNameU As String
    Dim strValidationRuleNameU As String

    Dim vsoValidationRuleSet As Visio.ValidationRuleSet

    strValidationRuleSetNameU = "Fault Tree Analysis"
    strValidationRuleNameU = "UngluedConnector"
    Set vsoValidationRuleSet = ActiveDocument.Validation.RuleSets(strValidationRuleSetNameU)
    
    ' Delete the rule.
    vsoValidationRuleSet.Rules(strRuleNameU).Delete
    
End Sub
```


