---
title: Application.ProjectResourceNew Event (Project)
ms.prod: project-server
api_name:
- Project.Application.ProjectResourceNew
ms.assetid: 9b030fbc-5cca-df10-f7a3-613d7ad70dc7
ms.date: 06/08/2017
---


# Application.ProjectResourceNew Event (Project)

Occurs before one or more resources is created.


## Syntax

 _expression_. **ProjectResourceNew**( ** _pj_**, ** _ID_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _pj_|Required|**Project**|The project in which the resource or resources are being created.|
| _ID_|Required|**Long**|ID of the new resource in the project.|

### Return Value

nothing


## Remarks

Project events do not occur when the project is embedded in another document or application.

The  **ProjectResourceNew** event doesn't occur during resource pool operations, when inserting or removing a subproject, or when changes have been made using a custom form.


