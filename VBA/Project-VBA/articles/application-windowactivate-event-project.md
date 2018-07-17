---
title: Application.WindowActivate Event (Project)
ms.prod: project-server
api_name:
- Project.Application.WindowActivate
ms.assetid: b54d0956-7eab-db5f-394a-5120bc111afd
ms.date: 06/08/2017
---


# Application.WindowActivate Event (Project)

Occurs when any window within Project is activated. The  **WindowActivate** event does not occur when the application window is activated.


## Syntax

 _expression_. **WindowActivate**( ** _activatedWindow_**, )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _activatedWindow_|Required|**Window**|The activated window.|

### Return Value

nothing


## Remarks

Project events do not occur when the project is embedded in another document or application.


