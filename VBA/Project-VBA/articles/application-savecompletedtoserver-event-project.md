---
title: Application.SaveCompletedToServer Event (Project)
ms.prod: project-server
api_name:
- Project.Application.SaveCompletedToServer
ms.assetid: 05ca27a0-a6cd-efbd-eff8-4f457c3de5c0
ms.date: 06/08/2017
---


# Application.SaveCompletedToServer Event (Project)

Occurs when Project Professional successfully puts the  **Project Save** job in the Project Server Queue.


## Syntax

 _expression_. **SaveCompletedToServer**( ** _bstrName_**, ** _bstrprojGuid_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _bstrName_|Required|**String**|Name of the project.|
| _bstrprojGuid_|Required|**String**|GUID of the project|

### Return Value

nothing


