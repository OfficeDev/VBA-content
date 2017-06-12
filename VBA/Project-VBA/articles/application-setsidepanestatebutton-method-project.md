---
title: Application.SetSidepaneStateButton Method (Project)
ms.prod: project-server
api_name:
- Project.Application.SetSidepaneStateButton
ms.assetid: 21603c44-d9f3-96b6-ee42-df17eb58287a
ms.date: 06/08/2017
---


# Application.SetSidepaneStateButton Method (Project)

Sets the state of the  **Toggle** button in the Project Guide. Deprecated in Project.


## Syntax

 _expression_. **SetSidepaneStateButton**( ** _DisplayState_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DisplayState_|Required|**Boolean**|**False** if the button should be in normal state, indicating the side pane is hidden. **True** if the button should be in depressed state, indicating the side pane is showing.|

## Remarks

The Project Guide is disabled by default in Project. Although you can create and display custom Project Guide pages, we recommend that you create task pane apps for new development.

The state of the button should be depressed when the side pane is showing, normal when the side pane is hidden.


