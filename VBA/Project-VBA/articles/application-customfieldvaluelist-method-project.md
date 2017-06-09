---
title: Application.CustomFieldValueList Method (Project)
keywords: vbapj.chm40
f1_keywords:
- vbapj.chm40
ms.prod: project-server
api_name:
- Project.Application.CustomFieldValueList
ms.assetid: 7365511c-6746-869b-f8e7-d4b87c5b8e70
ms.date: 06/08/2017
---


# Application.CustomFieldValueList Method (Project)

Sets options for a value list for a custom field.


## Syntax

 _expression_. **CustomFieldValueList**( ** _FieldID_**, ** _ListDefault_**, ** _DefaultValue_**, ** _RestrictToList_**, ** _AppendNew_**, ** _PromptOnNew_**, ** _DisplayOrder_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FieldID_|Required|**Long**|The custom field. Can be one of the  **[PjCustomField](pjcustomfield-enumeration-project.md)** constants.|
| _ListDefault_|Optional|**Boolean**|**True** if a value in the list functions as the default for the custom field.|
| _DefaultValue_|Optional|**String**|The item in the value list that is the default for the custom field. If  **ListDefault** is **False**, **DefaultValue** is ignored.|
| _RestrictToList_|Optional|**Boolean**|**True** if the only values allowed in the custom field are those from the value list.|
| _AppendNew_|Optional|**Boolean**|**True** if new values entered into the custom field are automatically added to the value list. If **RestrictToList** is **False**, **AppendNew** is ignored.|
| _PromptOnNew_|Optional|**Boolean**|**True** if the user is prompted to confirm that a new value is to be added to the list. If **AppendNew** is **False**, **PromptOnNew** is ignored.|
| _DisplayOrder_|Optional|**Long**|The order in which the items in a value list are displayed in the drop-down list for a cell. Can be one of the following  **PjListOrder** constants: **pjListOrderDefault**, **pjListOrderAscending**, or **pjListOrderDescending**.|

### Return Value

 **Boolean**


