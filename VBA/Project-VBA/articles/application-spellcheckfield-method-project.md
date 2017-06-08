---
title: Application.SpellCheckField Method (Project)
keywords: vbapj.chm2252
f1_keywords:
- vbapj.chm2252
ms.prod: project-server
api_name:
- Project.Application.SpellCheckField
ms.assetid: 4c5cc4c9-b947-c237-7f7e-0d703bd34352
ms.date: 06/08/2017
---


# Application.SpellCheckField Method (Project)

Checks the spelling of text custom fields.


## Syntax

 _expression_. **SpellCheckField**( ** _FieldName_**, ** _EnableSpellCheck_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FieldName_|Optional|**PjSpellingField**|One of the  **[PjSpellingField](pjspellingfield-enumeration-project.md)** enumeration values.|
| _EnableSpellCheck_|Optional|**Variant**|**True** if spell check is enabled; otherwise, **False**.|

### Return Value

 **Boolean**


## Remarks

To check spelling in the entire project, including text custom fields, use the  **[SpellingCheck](application-spellingcheck-method-project.md)** method. The **SpellingCheck** method is equivalent to the **Spelling** command on the **Project** tab of the Ribbon.


