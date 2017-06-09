---
title: Resource.EnterpriseTeamMember Method (Project)
ms.prod: project-server
api_name:
- Project.Resource.EnterpriseTeamMember
ms.assetid: a89acb10-02c3-0e2d-66b2-2d448514d919
ms.date: 06/08/2017
---


# Resource.EnterpriseTeamMember Method (Project)

Indicates whether the resource belongs to the project.  **True** if the resource is a member of the team for the specified project; otherwise **False**. Available in Project Professional only.


## Syntax

 _expression_. **EnterpriseTeamMember**( ** _Project_** )

 _expression_ A variable that represents a **Resource** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Project_|Required|**Object**|The  **Project** object against which the expression is checked. For example, **ActiveProject**.|

### Return Value

 **Boolean**


## Remarks

The  **EnterpriseTeamMember** method returns **False** for summary resource assignments, because the assignment or resource is from another project.

The  **EnterpriseTeamMember** method returns a trappable error (error code 1004) if the active view is not a Resource or Assignment view.


