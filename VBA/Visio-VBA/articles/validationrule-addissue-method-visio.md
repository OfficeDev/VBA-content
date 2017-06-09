---
title: ValidationRule.AddIssue Method (Visio)
keywords: vis_sdr.chm18462405
f1_keywords:
- vis_sdr.chm18462405
ms.prod: visio
api_name:
- Visio.ValidationRule.AddIssue
ms.assetid: 9ee6b555-a90a-c887-9869-ae2e307591f5
ms.date: 06/08/2017
---


# ValidationRule.AddIssue Method (Visio)

Creates a new validation issue that is based on the validation rule, and adds it to the document.


## Syntax

 _expression_ . **AddIssue**( **_[TargetPage]_** , **_[TargetShape]_** )

 _expression_ A variable that represents a **[ValidationRule](validationrule-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _TargetPage_|Optional| **[Page](page-object-visio.md)**|The page that has the issue. May be  **Nothing** .|
| _TargetShape_|Optional| **[Shape](shape-object-visio.md)**|The shape that has the issue. May be  **Nothing** .|

### Return Value

 **[ValidationIssue](validationissue-object-visio.md)**


## Remarks

 _TargetPage_ and _TargetShape_ identify the specific object that is associated with the issue. If the object that you pass for either parameter is not a valid object, or if it is inconsistent with the rule's target type, Microsoft Visio returns an Invalid Parameter error.

If you do not pass a value for the optional  _TargetShape_ parameter, the validation issue target is the page.

If you do not pass values for either of the optional parameters, the validation issue target is the document.


## Example

The following Visual Basic for Applications (VBA) example shows how to use the  **AddIssue** method to add a custom validation issue to a validation rule and associate it with a particular shape on a particular page.


```vb
' Add a custom issue to the vsoValidationRule validation rule and 
' associate it with shape vsoShape on page vsoPage.
Set vsoValidationIssue = vsoValidationRule.AddIssue(vsoPage, vsoShape)
```


