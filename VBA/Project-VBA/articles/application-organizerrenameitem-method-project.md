---
title: Application.OrganizerRenameItem Method (Project)
keywords: vbapj.chm130
f1_keywords:
- vbapj.chm130
ms.prod: project-server
api_name:
- Project.Application.OrganizerRenameItem
ms.assetid: 97ef4b63-a2fb-35ac-0a27-ebe8566fd28c
ms.date: 06/08/2017
---


# Application.OrganizerRenameItem Method (Project)

Renames an item in the Organizer.


## Syntax

 _expression_. **OrganizerRenameItem**( ** _Type_**, ** _Filename_**, ** _Name_**, ** _NewName_**, ** _Task_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Optional|**Long**|The type of item to rename. Can be one of the  **[PjOrganizer](pjorganizer-enumeration-project.md)** constants. The default value is **pjViews**.|
| _Filename_|Required|**String**|The name of the file containing the item to rename.|
| _Name_|Required|**String**|The name of the item to rename.|
| _NewName_|Required|**String**|The new name for the item specified by  **Name**.|
| _Task_|Optional|**Boolean**|**True** if the item applies to tasks. **False** if the item applies to resources. The default value is **True**.|

### Return Value

 **Boolean**


