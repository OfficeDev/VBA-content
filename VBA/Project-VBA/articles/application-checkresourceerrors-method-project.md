---
title: Application.CheckResourceErrors Method (Project)
keywords: vbapj.chm2258
f1_keywords:
- vbapj.chm2258
ms.prod: project-server
api_name:
- Project.Application.CheckResourceErrors
ms.assetid: 780cf9c8-078b-3707-f0e4-a468432c1ced
ms.date: 06/08/2017
---


# Application.CheckResourceErrors Method (Project)

Checks for errors when resources are imports to the enterprise, or when enterprise resource pool is saved.


## Syntax

 _expression_. **CheckResourceErrors**( ** _LocalRUID_**, ** _ResetImport_**, ** _CheckEnterprise_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _LocalRUID_|Optional|**Variant**|Local resource IDs that are to be checked for errors. (Example: "1, 5, 6, 7, 12".) If  **null**, all local resources are checked (unless CheckEnterprise is **True** ).|
| _ResetImport_|Optional|**Boolean**|Reset the import column to  **True** for the local resources that are being checked for errors.|
| _CheckEnterprise_|Optional|**Boolean**|If  **True**, check enterprise resources only.|

### Return Value

 **Boolean**


