---
title: Application.CheckTaskErrors Method (Project)
keywords: vbapj.chm2257
f1_keywords:
- vbapj.chm2257
ms.prod: project-server
api_name:
- Project.Application.CheckTaskErrors
ms.assetid: 7b361295-993a-13b2-b9bb-26f149e16e72
ms.date: 06/08/2017
---


# Application.CheckTaskErrors Method (Project)

Checks the task to ensure that required custom fields are filled and that the calendars have the enterprise calendars type. If the TaskID parameter is  **null**, all tasks are checked. .


## Syntax

 _expression_. **CheckTaskErrors**( ** _TaskID_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _TaskID_|Optional|**Variant**|TaskID for the task or  **null**.|

### Return Value

 **Boolean**


