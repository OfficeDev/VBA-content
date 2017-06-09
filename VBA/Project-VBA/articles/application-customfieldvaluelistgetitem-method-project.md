---
title: Application.CustomFieldValueListGetItem Method (Project)
keywords: vbapj.chm131200
f1_keywords:
- vbapj.chm131200
ms.prod: project-server
api_name:
- Project.Application.CustomFieldValueListGetItem
ms.assetid: 54ab8b15-374a-3c7a-ffe6-bc90b5d4561e
ms.date: 06/08/2017
---


# Application.CustomFieldValueListGetItem Method (Project)

Returns the value, description, or phonetic spelling of an item in the value list for a custom field.


## Syntax

 _expression_. **CustomFieldValueListGetItem**( ** _FieldID_**, ** _Item_**, ** _Index_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FieldID_|Required|**Long**|The custom field. Can be one of the  **[PjCustomField](pjcustomfield-enumeration-project.md)** constants.|
| _Item_|Required|**Long**|The information to return. Can be one of the following  **PjValueListItem** constants: **pjValueListValue**, **pjValueListDescription**, or **pjValueListPhonetics**. The default value is **pjValueListValue**.|
| _Index_|Required|**Long**|The row number of the value list item for which to return the information specified with Item.|

### Return Value

 **String**


