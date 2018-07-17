---
title: Application.BoxDataTemplate Method (Project)
keywords: vbapj.chm2391
f1_keywords:
- vbapj.chm2391
ms.prod: project-server
api_name:
- Project.Application.BoxDataTemplate
ms.assetid: ce3530d5-6218-b0db-a890-9a80bca5e3db
ms.date: 06/08/2017
---


# Application.BoxDataTemplate Method (Project)

Creates, copies, renames, or deletes a data template for a Network Diagram view.


## Syntax

 _expression_. **BoxDataTemplate**( ** _Name_**, ** _Action_**, ** _NewName_**, ** _Overwrite_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**|The name of the template to edit, create, copy or delete.|
| _action_|Required|**Long**|The operation to perform on the template. Can be one of the following  **[PjDataTemplate](pjdatatemplate-enumeration-project.md)** constants: **pjDataTemplateCopy**, **pjDataTemplateDelete**, **pjDataTemplateNew**, or **pjDataTemplateRename**.|
| _NewName_|Optional|**String**|Required when specifying a new name for an existing data template ( **action** is **pjDataTemplateNew** ) or naming a copied data template ( **action** is **pjDataTemplateCopy** ). If **action** is **pjDataTemplateRename** or **pjDataTemplateDelete**, **NewName** is ignored.|
| _Overwrite_|Optional|**Boolean**|**True** if an existing template should be replaced with one of the same name. If **action** is **pjDataTemplateRename** or **pjDataTemplateDelete**, **Overwrite** is ignored. The default value is **False**.|

### Return Value

 **Boolean**


