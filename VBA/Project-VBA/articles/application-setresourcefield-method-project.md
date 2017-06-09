---
title: Application.SetResourceField Method (Project)
keywords: vbapj.chm5
f1_keywords:
- vbapj.chm5
ms.prod: project-server
api_name:
- Project.Application.SetResourceField
ms.assetid: fbf71bbe-86cc-c53c-a0c3-0df288e2b480
ms.date: 06/08/2017
---


# Application.SetResourceField Method (Project)

Sets the value of a resource field.


## Syntax

 _expression_. **SetResourceField**( ** _Field_**, ** _Value_**, ** _AllSelectedResources_**, ** _Create_**, ** _ResourceID_**, ** _ProjectName_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Field_|Required|**String**|The name of the resource field to set.|
| _Value_|Required|**String**|The value of the resource field.|
| _AllSelectedResources_|Optional|**Boolean**|**True** if the value of the field is set for all selected resources. **False** if the value is set for the active resource. The default value is **False**.|
| _Create_|Optional|**Boolean**|**True** if Project should create a new resource if the active cell is on an empty row. The default value is **True**.|
| _ResourceID_|Optional|**Long**|The identification number of the resource containing the field to set. If AllSelectedResources is  **True**, ResourceID is ignored.|
| _ProjectName_|Optional|**String**|If the active project is a consolidated project, specifies the name of the project for the resource specified by ResourceID . If ResourceID is not specified, ProjectName is ignored. The default value is the name of the active project.|

### Return Value

 **Boolean**


