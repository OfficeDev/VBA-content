---
title: Application.SetResourceFieldByID Method (Project)
keywords: vbapj.chm96
f1_keywords:
- vbapj.chm96
ms.prod: project-server
api_name:
- Project.Application.SetResourceFieldByID
ms.assetid: 1309ee61-6b66-db45-ed69-b0b3dd9b8dda
ms.date: 06/08/2017
---


# Application.SetResourceFieldByID Method (Project)

Sets the value of a resource field specified by the field identification number.


## Syntax

 _expression_. **SetResourceFieldByID**( ** _FieldID_**, ** _Value_**, ** _AllSelectedResources_**, ** _Create_**, ** _ResourceID_**, ** _ProjectName_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FieldID_|Required|**PjField**|Field identification number. Can be one of the resource fields specified by a  **[PjField](pjfield-enumeration-project.md)** constant or a number returned by the **[FieldNameToFieldConstant](application-fieldnametofieldconstant-method-project.md)** method.|
| _Value_|Required|**String**|The value of the resource field.|
| _AllSelectedResources_|Optional|**Boolean**|**True** if the value of the field is set for all selected resources. **False** if the value is set for the active resource. The default value is **False**.|
| _Create_|Optional|**Boolean**|**True** if Project should create a resource if the active cell is on an empty row. The default value is **True**.|
| _ResourceID_|Optional|**Long**|The identification number of the resource containing the field to set. If AllSelectedResources is  **True**, ResourceID is ignored.|
| _ProjectName_|Optional|**String**|If the active project is a consolidated project, specifies the name of the project for the resource specified by  _ResourceID_. If  _ResourceID_ is not specified, _ProjectName_ is ignored. The default value is the name of the active project.|

### Return Value

 **Boolean**


## Remarks

To set a resource field by name, use the  **[SetResourceField](application-setresourcefield-method-project.md)** method.


