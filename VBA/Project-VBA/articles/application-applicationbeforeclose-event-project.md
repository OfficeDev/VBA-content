---
title: Application.ApplicationBeforeClose Event (Project)
ms.prod: project-server
api_name:
- Project.Application.ApplicationBeforeClose
ms.assetid: 9523a793-b4c1-fd79-303e-b167d7f80025
ms.date: 06/08/2017
---


# Application.ApplicationBeforeClose Event (Project)

Occurs before Project exits.


## Syntax

 _expression_. **ApplicationBeforeClose**( ** _Info_**, )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Info_|Required|**EventInfo**|**EventInfo.Cancel** is **False** when the event occurs. If the event procedure sets this argument to **True**, Project does not close when the procedure is finished.|

### Return Value

nothing


## Remarks

Project events do not occur when the project is embedded in another document or application.


