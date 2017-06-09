---
title: Application.OrganizerDeleteItem Method (Project)
keywords: vbapj.chm128
f1_keywords:
- vbapj.chm128
ms.prod: project-server
api_name:
- Project.Application.OrganizerDeleteItem
ms.assetid: 7c243672-0e31-e224-eadd-3545f7efcde4
ms.date: 06/08/2017
---


# Application.OrganizerDeleteItem Method (Project)

Deletes an item from the Organizer.


## Syntax

 _expression_. **OrganizerDeleteItem**( ** _Type_**, ** _Filename_**, ** _Name_**, ** _Task_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Optional|**Long**|The type of item to delete. Can be one of the  **[PjOrganizer](pjorganizer-enumeration-project.md)** constants. The default value is **pjViews**.|
| _Filename_|Required|**String**|The name of the file containing the item to delete.|
| _Name_|Required|**String**|The name of the item to delete.|
| _Task_|Optional|**Boolean**|**True** if the item applies to tasks. **False** if the item applies to resources. The default value is **True**.|

### Return Value

 **Boolean**


