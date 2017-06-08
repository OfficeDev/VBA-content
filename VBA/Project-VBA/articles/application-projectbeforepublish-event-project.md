---
title: Application.ProjectBeforePublish Event (Project)
ms.prod: project-server
api_name:
- Project.Application.ProjectBeforePublish
ms.assetid: 5778ec6c-a8c0-0a05-145c-c9ad6132bf87
ms.date: 06/08/2017
---


# Application.ProjectBeforePublish Event (Project)

Occurs before a  **Publish** operation is placed on the server queue. The **ProjectBeforePublish** event can be cancelled. Project Professional only.


## Syntax

 _expression_. **ProjectBeforePublish**( ** _pj_**, ** _Cancel_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _pj_|Required|**Project**|Project object.|
| _Cancel_|Required|**Boolean**|**True** to cancel the **Publish** job.|

### Return Value

Nothing


## Remarks

The  **ProjectBeforePublish** event is commonly used to determine whether certain conditions are satisfied and to cancel publishing if the conditions are not met.


