---
title: Application.ProjectBeforeResourceDelete Event (Project)
ms.prod: project-server
api_name:
- Project.Application.ProjectBeforeResourceDelete
ms.assetid: aadef12e-57dc-210e-d29a-54f79d1c1abd
ms.date: 06/08/2017
---


# Application.ProjectBeforeResourceDelete Event (Project)

Occurs before a resource is deleted.


## Syntax

 _expression_. **ProjectBeforeResourceDelete**( ** _res_**, ** _Cancel_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _res_|Required|**Resource**| The resource that is being deleted.|
| _Cancel_|Required|**Boolean**|**False** when the event occurs. If the event procedure sets this argument to **True**, the resource is not deleted.|

### Return Value

nothing


## Remarks

Project events do not occur when the project is embedded in another document or application.

The  **ProjectBeforeResourceDelete** event doesn't occur when changes have been made using a custom form.


