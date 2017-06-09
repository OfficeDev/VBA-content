---
title: Application.EnterpriseResourceGet Method (Project)
keywords: vbapj.chm2094
f1_keywords:
- vbapj.chm2094
ms.prod: project-server
api_name:
- Project.Application.EnterpriseResourceGet
ms.assetid: c1e29298-7859-28c4-edbf-917acdd8aecd
ms.date: 06/08/2017
---


# Application.EnterpriseResourceGet Method (Project)

Adds a single resource from the enterprise resource pool to the active project. Available in Project Professional only.


## Syntax

 _expression_. **EnterpriseResourceGet**( ** _EUID_**, ** _RUID_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _EUID_|Optional|**Variant**|The unique ID of the enterprise resource; required if there is more than one resource.|
| _RUID_|Optional|**Variant**|The unique ID that will be assigned to the resource in the active project. If omitted, the next valid resource UID is assigned.|

### Return Value

 **Boolean**


