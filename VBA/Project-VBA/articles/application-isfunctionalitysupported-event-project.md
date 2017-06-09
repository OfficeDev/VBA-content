---
title: Application.IsFunctionalitySupported Event (Project)
ms.prod: project-server
api_name:
- Project.Application.IsFunctionalitySupported
ms.assetid: f6462a3b-5a36-3b2e-79bd-78cce567aed8
ms.date: 06/08/2017
---


# Application.IsFunctionalitySupported Event (Project)

Occurs after the  **LoadWebBrowserControl** method is called with the third parameter ( _FunctionalityName_) set.


## Syntax

 _expression_. **IsFunctionalitySupported**( ** _bstrFunctionality_**, ** _Info_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _bstrFunctionality_|Required|**String**|Name of the functionality.|
| _Info_|Required|**EventInfo**|EventInfo object.|

### Return Value

nothing


