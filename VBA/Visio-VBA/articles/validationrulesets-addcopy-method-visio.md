---
title: ValidationRuleSets.AddCopy Method (Visio)
keywords: vis_sdr.chm18160420
f1_keywords:
- vis_sdr.chm18160420
ms.prod: visio
api_name:
- Visio.ValidationRuleSets.AddCopy
ms.assetid: a9510a97-7a85-3e68-6493-2a43840ef934
ms.date: 06/08/2017
---


# ValidationRuleSets.AddCopy Method (Visio)

Adds a copy of an existing  **[ValidationRuleSet](validationruleset-object-visio.md)** object to the **[ValidationRuleSets](validationrulesets-object-visio.md)** collection of the document.


## Syntax

 _expression_ . **AddCopy**( **_RuleSet_** , **_NameU_** )

 _expression_ A variable that represents a **ValidationRuleSet** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _RuleSet_|Required| **ValidationRuleSet**|The existing rule set to copy and add to the collection of rule sets.|
| _NameU_|Optional| **String**|The universal name to assign to the new validation rule set.|

### Return Value

 **ValidationRuleSet**


## Remarks

If you pass a value for the optional  _NameU_ parameter, both the **[Name](validationruleset-name-property-visio.md)** and **[NameU](validationruleset-nameu-property-visio.md)** properties of the new rule set are assigned the value. If you do not pass a value, Microsoft Visio assigns the new rule set the local and universal name of the existing rule set. In that case, if you copy a rule set within a document, Visio overwrites the existing rule set. However, if you copy a rule set to another document, Visio adds a new rule set to the other document and leaves the existing rule set unchanged.

Similarly, if the value that you pass for the  _NameU_ matches the universal name of an existing rule set in the document, Visio overwrites the existing rule set.

If the  _NameU_ parameter is not a valid string, Visio returns an Invalid Parameter error.


