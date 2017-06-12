---
title: Assignment.EnterpriseTeamMember Method (Project)
ms.prod: project-server
api_name:
- Project.Assignment.EnterpriseTeamMember
ms.assetid: 706a7f8b-b545-7398-7c09-f29f6b8d225d
ms.date: 06/08/2017
---


# Assignment.EnterpriseTeamMember Method (Project)

Indicates whether the specified assignment belongs to the project.  **True** if the assignment belongs to the specified project; otherwise **False**. Available in Project Professional only.


## Syntax

 _expression_. **EnterpriseTeamMember**( ** _Project_** )

 _expression_ A variable that represents an **Assignment** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Project_|Required|**Object**|The  **Project** object against which the expression is checked. For example, **ActiveProject**.|

### Return Value

 **Boolean**


## Remarks

The  **EnterpriseTeamMember** method returns **False** for summary resource assignments, because the assignment or resource is from another project.

The  **EnterpriseTeamMember** method returns a trappable error (error code 1004) if the active view is not a Resource or Assignment view.


