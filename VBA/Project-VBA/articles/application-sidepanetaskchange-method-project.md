---
title: Application.SidepaneTaskChange Method (Project)
keywords: vbapj.chm53
f1_keywords:
- vbapj.chm53
ms.prod: project-server
api_name:
- Project.Application.SidepaneTaskChange
ms.assetid: 277a9242-b098-8f69-44b8-668175867b42
ms.date: 06/08/2017
---


# Application.SidepaneTaskChange Method (Project)

Changes the side pane that is displayed in  **Project Guide**.


## Syntax

 _expression_. **SidepaneTaskChange**( ** _ID_**, ** _IsGoalArea_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ID_|Required|**Integer**|The ID number of the side pane in the  **Project Guide**.|
| _IsGoalArea_|Optional|**Boolean**|**True** if trying to change to a different goal area in the **Project Guide**.  **False** if trying to change to a different **Project Guide** task.|

### Return Value

 **Boolean**


## Remarks

The  **SidepaneTaskChange** method only has an effect when the **Project Guide** is shown.


