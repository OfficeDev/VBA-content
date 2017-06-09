---
title: Application.EnterpriseResourcesOpen Method (Project)
keywords: vbapj.chm2088
f1_keywords:
- vbapj.chm2088
ms.prod: project-server
api_name:
- Project.Application.EnterpriseResourcesOpen
ms.assetid: 343b5391-2a28-043d-8ee9-34c71003126c
ms.date: 06/08/2017
---


# Application.EnterpriseResourcesOpen Method (Project)

Opens the pool of enterprise resources for viewing in a temporary project. Available in Project Professional only.


## Syntax

 _expression_. **EnterpriseResourcesOpen**( ** _EUID_**, ** _OpenType_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _EUID_|Optional|**Variant**|A comma-delimited list of unique IDs of the enterprise resource.|
| _OpenType_|Optional|**Long**|Specifies how the enterprise resources are to be checked out. Can be one of the following  **PjCheckOutType** constants: **pjReadOnly** or **pjReadWrite**. The default value is **pjReadWrite**.|

### Return Value

 **Boolean**


