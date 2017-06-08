---
title: Application.DDEExecute Method (Project)
keywords: vbapj.chm1202
f1_keywords:
- vbapj.chm1202
ms.prod: project-server
api_name:
- Project.Application.DDEExecute
ms.assetid: 307b1373-309a-1ecf-6899-fd64e663e4f9
ms.date: 06/08/2017
---


# Application.DDEExecute Method (Project)

Performs actions or runs commands in another application through dynamic data exchange (DDE).


## Syntax

 _expression_. **DDEExecute**( ** _Command_**, ** _TimeOut_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Command_|Required|**String**|The command to carry out in another application.|
| _TimeOut_|Optional|**Variant**|The number of seconds to wait for the other application to execute before proceeding. The default value is 5.|

### Return Value

 **Boolean**


## Remarks

If your macro displays a dialog box in another application, you may need to increase the default value for Timeout.


