---
title: Application.SetTaskFieldByID Method (Project)
keywords: vbapj.chm95
f1_keywords:
- vbapj.chm95
ms.prod: project-server
api_name:
- Project.Application.SetTaskFieldByID
ms.assetid: b4c74d96-d25b-707e-15f1-5e7f05363360
ms.date: 06/08/2017
---


# Application.SetTaskFieldByID Method (Project)

Sets the value of a task field specified by the field identification number.


## Syntax

 _expression_. **SetTaskFieldByID**( ** _FieldID_**, ** _Value_**, ** _AllSelectedTasks_**, ** _Create_**, ** _TaskID_**, ** _ProjectName_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FieldID_|Required|**PjField**|Field identification number. Can be one of the task fields specified by a  **[PjField](pjfield-enumeration-project.md)** constant or a number returned by the **[FieldNameToFieldConstant](application-fieldnametofieldconstant-method-project.md)** method.|
| _Value_|Required|**String**|The value of the task field.|
| _AllSelectedTasks_|Optional|**Variant**|**True** if the value of the field is set for all selected tasks. **False** if the value is set for the active task. The default value is **False**.|
| _Create_|Optional|**Variant**|**True** if Project creates a task when the active cell is on an empty row. The default value is **True**.|
| _TaskID_|Optional|**Variant**|The identification number of the task containing the field to set. If  _AllSelectedResources_ is **True**, _TaskID_ is ignored.|
| _ProjectName_|Optional|**Variant**|If the active project is a consolidated project, specifies the name of the project for the task specified by  _TaskID_. If  _TaskID_ is not specified, _ProjectName_ is ignored. The default value is the name of the active project.|

### Return Value

 **Boolean**


## Remarks

To set a task field by name, use the  **[SetTaskField](application-settaskfield-method-project.md)** method.


