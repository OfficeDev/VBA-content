---
title: Validation.RuleSets Property (Visio)
keywords: vis_sdr.chm18062715
f1_keywords:
- vis_sdr.chm18062715
ms.prod: visio
api_name:
- Visio.Validation.RuleSets
ms.assetid: cb75f7e0-f92c-86a9-3aee-21e1b0a4b16a
ms.date: 06/08/2017
---


# Validation.RuleSets Property (Visio)

Returns the collection of all the validation rule sets in the document. Read-only.


## Syntax

 _expression_ . **RuleSets**

 _expression_ A variable that represents a **[Validation](validation-object-visio.md)** object.


### Return Value

 **[ValidationRuleSets](validationrulesets-object-visio.md)**


## Example

The following sample code is based on code provided by: [David Parker](http://www.bvisual.net)

The following Visual Basic for Applications (VBA) example shows how to use the  **RuleSets** property to get the names of all the validation rule sets in the active document and print those names in the **Immediate** window.




```vb
Public Sub RuleSets_Example()

    Dim vsoDocument As Visio.Document
    Dim vsoRuleSet As Visio.ValidationRuleSet
    Dim vsoValidationRule As Visio.ValidationRule
    
    Set vsoDocument = Visio.ActiveDocument
    
    For Each vsoRuleSet In vsoDocument.Validation.RuleSets
    
        If vsoRuleSet.Enabled Then
            
            Debug.Print vsoRuleSet.NameU
            
        End If
    Next 
       
End Sub
```


