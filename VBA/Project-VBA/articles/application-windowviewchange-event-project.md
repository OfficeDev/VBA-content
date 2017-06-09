---
title: Application.WindowViewChange Event (Project)
ms.prod: project-server
api_name:
- Project.Application.WindowViewChange
ms.assetid: e6a5f884-5bb9-f975-9237-25996b436589
ms.date: 06/08/2017
---


# Application.WindowViewChange Event (Project)

Occurs after the top pane view is changed within a project window. The  **WindowViewChange** event returns a success argument that tells whether the view change action was successful.


## Syntax

 _expression_. **WindowViewChange**( ** _Window_**, ** _prevView_**, ** _newView_**, ** _success_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Window_|Required|**Window**|The window where the view change occurs.|
| _prevView_|Required|**View**|The previous topic pane view the user was in before the view change occurred. If the user was not in a project view before applying the current view, the prevView argument returns null.|
| _newView_|Required|**View**|The new top pane view the user has now applied.|
| _success_|Required|**Boolean**|**True** if the view change action succeeded.|

### Return Value

nothing


## Remarks

Project events do not occur when the project is embedded in another document or application.


