---
title: Application.CustomFieldRename Method (Project)
keywords: vbapj.chm2378
f1_keywords:
- vbapj.chm2378
ms.prod: project-server
api_name:
- Project.Application.CustomFieldRename
ms.assetid: 0ca77914-1881-eee5-a8ec-7b47c6464969
ms.date: 06/08/2017
---


# Application.CustomFieldRename Method (Project)

Defines a friendly name for a custom field.


## Syntax

 _expression_. **CustomFieldRename**( ** _FieldID_**, ** _NewName_**, ** _Phonetic_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FieldID_|Required|**Long**|The field to be renamed. Can be one of the  **[PjCustomField](pjcustomfield-enumeration-project.md)** constants.|
| _NewName_|Optional|**String**|The friendly name for the custom field. A value of  **Null** removes the friendly name.|
| _Phonetic_|Optional|**String**|The phonetic equivalent of the friendly name. The Phonetic argument is ignored unless the Japanese version of Project is used.|

### Return Value

 **Boolean**


