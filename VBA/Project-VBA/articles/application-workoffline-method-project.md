---
title: Application.WorkOffline Method (Project)
keywords: vbapj.chm2283
f1_keywords:
- vbapj.chm2283
ms.prod: project-server
api_name:
- Project.Application.WorkOffline
ms.assetid: 65a38e80-f311-eb19-359a-da9f1022be71
ms.date: 06/08/2017
---


# Application.WorkOffline Method (Project)

Opens or closes the connection to Project Server. 


## Syntax

 _expression_. **WorkOffline**( ** _fOffline_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _fOffline_|Optional|**Boolean**|**True** closes the connection to Project Server. **False** opens the connection to Project Server.|

### Return Value

 **Boolean**


## Remarks

Available in Project Professional only. If Project is started with an offline account, the WorkOffline method results in a run-time error 1100: "The method is not available in this situation."


