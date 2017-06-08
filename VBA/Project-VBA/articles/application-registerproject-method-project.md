---
title: Application.RegisterProject Method (Project)
keywords: vbapj.chm131250
f1_keywords:
- vbapj.chm131250
ms.prod: project-server
api_name:
- Project.Application.RegisterProject
ms.assetid: 66cc4443-2adc-ff66-976e-da52c6d4f7ff
ms.date: 06/08/2017
---


# Application.RegisterProject Method (Project)

Registers the active project on Project Server.


## Syntax

 _expression_. **RegisterProject**( ** _WaitForCompletion_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _WaitForCompletion_|Required|**Boolean**|**True** if Project waits until the registration is complete before returning notification that the operation was successful or returning an error code if the operation failed. The default value is **False**.|

### Return Value

 **Long**


## Remarks

The  **RegisterProject** method is available only in Project Professional.


