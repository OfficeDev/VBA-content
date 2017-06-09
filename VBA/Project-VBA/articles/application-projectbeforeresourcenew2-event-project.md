---
title: Application.ProjectBeforeResourceNew2 Event (Project)
ms.prod: project-server
api_name:
- Project.Application.ProjectBeforeResourceNew2
ms.assetid: 24c28eac-946b-80fb-5dcb-8b9ef499b547
ms.date: 06/08/2017
---


# Application.ProjectBeforeResourceNew2 Event (Project)

Occurs before one or more resources are created. Uses the  **EventInfo** object parameter.


## Syntax

 _expression_. **ProjectBeforeResourceNew2**( ** _pj_**, ** _Info_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _pj_|Required|**Project**|The project in which the resource or resources are being created.|
| _Info_|Required|**EventInfo**|EventInfo.Cancel is  **False** when the event occurs. If the event procedure sets this argument to **True**, the new resource or resources are not created.|

### Return Value

nothing


## Remarks

Project events do not occur when the project is embedded in another document or application.

The  **ProjectBeforeResourceNew2** event doesn't occur during resource pool operations, when inserting or removing a subproject, or when changes have been made using a custom form.


