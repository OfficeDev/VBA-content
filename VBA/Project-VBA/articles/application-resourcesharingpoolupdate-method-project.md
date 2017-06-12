---
title: Application.ResourceSharingPoolUpdate Method (Project)
keywords: vbapj.chm248
f1_keywords:
- vbapj.chm248
ms.prod: project-server
api_name:
- Project.Application.ResourceSharingPoolUpdate
ms.assetid: 1ebcf06f-fce3-7403-2adb-56f60ab73259
ms.date: 06/08/2017
---


# Application.ResourceSharingPoolUpdate Method (Project)

Synchronizes the information in the sharer project with the information in the local resource pool project.


## Syntax

 _expression_. **ResourceSharingPoolUpdate**( ** _allSharers_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _allSharers_|Optional|**Boolean**|**True** if the information from all open sharers is updated in the pool. **False** if only the information from sharers in the active project is updated in the pool. If **AllSharers** is omitted and only one sharer is open, that information is updated in the pool; otherwise, the user is prompted to specify whether all open sharers or just those in the active project should be updated in the pool.|

### Return Value

 **Boolean**


## Remarks




 **Note**  Project Professional can share local resources only when not logged on Project Server. If Project Professional is using a Project Server profile, local resource sharing is unavailable.


## Example

In the following example, the project that contains the resources to share is named SharedResourcePool.mpp. If the active project is named Sharer.mpp, the code enables Sharer.mpp to synchronize with any changes in resources from SharedResourcePool.mpp. Both projects must be open.


```vb
Application.ResourceSharingPoolUpdate
```


