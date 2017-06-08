---
title: Application.CustomFieldSetFormula Method (Project)
keywords: vbapj.chm36
f1_keywords:
- vbapj.chm36
ms.prod: project-server
api_name:
- Project.Application.CustomFieldSetFormula
ms.assetid: d6d5a5d5-c948-07c9-3f5e-b4607df6538c
ms.date: 06/08/2017
---


# Application.CustomFieldSetFormula Method (Project)

Specifies a formula to use when assigning a value to a custom field.


## Syntax

 _expression_. **CustomFieldSetFormula**( ** _FieldID_**, ** _Formula_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FieldID_|Required|**Long**|The custom field. Can be one of the  **[PjCustomField](pjcustomfield-enumeration-project.md)** constants.|
| _Formula_|Optional|**String**|The formula to use to assign a value for the custom field. The value specified with  **Formula** functions as the right side of an equation that the field specified with **FieldID** should equal. To specify a field as part of the formula, enclose the field name in brackets, as in "[Actual Cost] * 2". If a macro will be run in more than one language, any field specified in **Formula** must use the name localized for each language.|

### Return Value

 **Boolean**


