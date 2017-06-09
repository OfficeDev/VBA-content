---
title: Application.WindowBeforeViewChange Event (Project)
ms.prod: project-server
api_name:
- Project.Application.WindowBeforeViewChange
ms.assetid: c3eb450d-2a74-6ae1-175c-1d61c90b22ca
ms.date: 06/08/2017
---


# Application.WindowBeforeViewChange Event (Project)

Occurs when the top pane view is changed within a window in Project.


## Syntax

 _expression_. **WindowBeforeViewChange**( ** _Window_**, ** _prevView_**, ** _newView_**, ** _projectHasViewWindow_**, ** _Info_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Window_|Required|**Window**|The window where the view change occurs.|
| _prevView_|Required|**View**|The previous view (top pane) the user is in. If the user was not in a project view before applying the current view, this value will return  **Null**.|
| _newView_|Required|**View**|The new view (top pane) to which the user is trying to change.|
| _projectHasViewWindow_|Required|**Boolean**|True if the Project  **View Bar** is currently visible.|
| _Info_|Required|**EventInfo**|EventInfo.Cancel is  **False** when the event occurs. If the event procedure sets this argument to **True**, the value for the field specified with Field is not changed.|

### Return Value

nothing


## Remarks

Project events do not occur when the project is embedded in another document or application.


