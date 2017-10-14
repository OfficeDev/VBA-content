---
title: Application.DDEInitiate Method (Project)
keywords: vbapj.chm1201
f1_keywords:
- vbapj.chm1201
ms.prod: project-server
api_name:
- Project.Application.DDEInitiate
ms.assetid: a517c66f-4bec-9bec-270c-2053bc733145
ms.date: 06/08/2017
---


# Application.DDEInitiate Method (Project)

Opens a dynamic data exchange (DDE) channel to an application.


## Syntax

 _expression_. **DDEInitiate**( ** _App_**, ** _Topic_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _App_|Required|**String**| The name of the application to which you want to send commands.|
| _Topic_|Required|**String**|A document in the application to which you want to send commands.|

### Return Value

 **Boolean**


