---
title: Validation.Validate Method (Visio)
keywords: vis_sdr.chm18062725
f1_keywords:
- vis_sdr.chm18062725
ms.prod: visio
api_name:
- Visio.Validation.Validate
ms.assetid: 9e8b8bcd-674e-c7ac-543c-027ed02519cd
ms.date: 06/08/2017
---


# Validation.Validate Method (Visio)

Validates the specified validation rule set.


## Syntax

 _expression_ . **Validate**( **_RuleSet_** , **_Flags_** )

 _expression_ A variable that represents a **[Validation](validation-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _RuleSet_|Optional| **[ValidationRuleSet](validationruleset-object-visio.md)**|The rule set to validate across the entire document. |
| _Flags_|Optional| **[VisValidationFlags](visvalidationflags-enumeration-visio.md)**|Whether to open the  **Issues** window after validation.|

### Return Value

 **Nothing**


## Remarks

To validate all rule sets active in the document, pass  **Nothing** for _RuleSet_ .

 _Flags_ must be one of the following **VisValidationFlags** constants.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visValidationDefault**|0|Validate document and open  **Issues** window. The default.|
| **visValidationNoOpenWindow**|1|Validate document but do not open  **Issues** window.|
If you do not set the optional  _Flags_ parameter, Microsoft Visio applies the default behavior ( **visValidationDefault** ).

When you call the  **Validate** method, Microsoft Visio checks whether the rule set is active before evaluating it. Visio does not display message boxes during the evalution, except to notify you if, when _Flags_ is set to **visValidationDefault** , it finds no errors; and it displays the progress bar only if **Application.ShowProgressBars** is **True** .


## Example

The following Visual Basic for Applications (VBA) example shows how to use the  **Validate** method to validate the active document.


```vb
' Validate the document.
Call Visio.ActiveDocument.Validation.Validate(,Visio.VisValidationFlags.visValidationDefault)
```


