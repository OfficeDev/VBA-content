---
title: Application.GroupMaintainHierarchy Method (Project)
keywords: vbapj.chm2296
f1_keywords:
- vbapj.chm2296
ms.prod: project-server
api_name:
- Project.Application.GroupMaintainHierarchy
ms.assetid: 63f5763a-0ca3-d25b-06ac-03e52cdcf6e2
ms.date: 06/08/2017
---


# Application.GroupMaintainHierarchy Method (Project)

Shows or hides item hierarchy in task views or resource views where a group is applied.


## Syntax

 _expression_. **GroupMaintainHierarchy**( ** _On_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _On_|Required|**Boolean**|**True** if hierarchy in the current group is maintained. **False** if hierarchy is not maintained.|

### Return Value

 **Boolean**


## Remarks

The  **GroupMaintainHierarchy** method corresponds to the following command on the ribbon: On the **View** tab, click the **Group by** drop-down list in the **Data** group, and then click **Maintain Hierarchy in Current Group**.

For example, if tasks are grouped by the Critical group, the  `GroupMaintainHierarchy True` command shows the summary tasks in the **Critical: No** and **Critical: Yes** groups. The `GroupMaintainHierarchy False` command hides summary tasks in the groups. If no group is applied to the view, **GroupMaintainHierarchy** has no effect.


