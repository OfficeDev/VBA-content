---
title: ValidationRule.RuleSet Property (Visio)
keywords: vis_sdr.chm18462670
f1_keywords:
- vis_sdr.chm18462670
ms.prod: visio
api_name:
- Visio.ValidationRule.RuleSet
ms.assetid: 0152d440-b476-fdbc-b6d1-8b0aa29e841a
ms.date: 06/08/2017
---


# ValidationRule.RuleSet Property (Visio)

Returns the  **[ValidationRuleSet](validationruleset-object-visio.md)** object that contains the specified validation rule. Read-only.


## Syntax

 _expression_ . **RuleSet**

 _expression_ A variable that represents a **[ValidationRule](validationrule-object-visio.md)** object.


### Return Value

 **ValidationRuleSet**


## Example

The following Visual Basic for Applications (VBA) example shows how to use the  **RuleSet** property to selectively delete validation issues that belong to a particular rule set.


```vb
Set vsoDocument = Visio.ActiveDocument 
Set vsoIssues = vsoDocument.Validation.Issues
intIssueTotal = vsoIssues.Count
intIssueNumber = 1

' Iterate through the validation issues.
 For intCurrentIssue = 1 To intIssueTotal
      Set vsoIssue = vsoDocument.Validation.Issues(intIssueNumber)
      
     ' Delete the issues that belong to the vsoValidationRuleSet rule set.
     If vsoIssue.Rule.RuleSet Is vsoValidationRuleSet Then
         vsoIssue.Delete
     Else
        intIssueNumber = intIssueNumber + 1
     End If
     
 Next intCurrentIssue
```


