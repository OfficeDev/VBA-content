---
title: Application.SecondaryViewChange Event (Project)
ms.prod: project-server
api_name:
- Project.Application.SecondaryViewChange
ms.assetid: f0f3f81b-c75f-79ee-db8b-6bdd32a3702f
ms.date: 06/08/2017
---


# Application.SecondaryViewChange Event (Project)

Event occurs when a secondary view pane changes within a project window.


## Syntax

 _expression_. **SecondaryViewChange**( ** _Window_**, ** _prevView_**, ** _newView_**, ** _success_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Window_|Required|**Window**|The name of the Project file.|
| _prevView_|Required|**View**|The name of the previous topic pane view before the view change occurred. If the user was not in a project view before applying the current view, the prevView argument returns  **null**.|
| _newView_|Required|**View**|The name of the new top pane view that the user applied. |
| _success_|Required|**Boolean**|**True** if the view change action succeeded.|

### Return Value

nothing


