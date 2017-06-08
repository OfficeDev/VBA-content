---
title: Application.ResourceSharingPoolAction Method (Project)
keywords: vbapj.chm2083
f1_keywords:
- vbapj.chm2083
ms.prod: project-server
api_name:
- Project.Application.ResourceSharingPoolAction
ms.assetid: 0406765b-b6d7-ad6b-c1c2-51bb55591e69
ms.date: 06/08/2017
---


# Application.ResourceSharingPoolAction Method (Project)

Performs the specified action on a local resource pool.


## Syntax

 _expression_. **ResourceSharingPoolAction**( ** _Action_**, ** _Filename_**, ** _ReadOnly_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _action_|Required|**Long**|The actions to perform on the resource pool. Can be one of the  **[PjPoolAction](pjpoolaction-enumeration-project.md)** constants.|
| _Filename_|Optional|**String**|The file name of the resource pool on which to perform the action.|
| _ReadOnly_|Optional|**Boolean**|**True** if the files specified with **FileName** are opened read-only.|

### Return Value

 **Boolean**


## Remarks




 **Note**  Project Professional can share local resources only when not logged on Project Server. If Project Professional is using a Project Server profile, local resource sharing is unavailable.


