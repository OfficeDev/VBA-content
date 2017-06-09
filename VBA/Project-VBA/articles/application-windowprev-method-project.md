---
title: Application.WindowPrev Method (Project)
keywords: vbapj.chm2006
f1_keywords:
- vbapj.chm2006
ms.prod: project-server
api_name:
- Project.Application.WindowPrev
ms.assetid: f95cf733-fc5c-e454-55b6-11f704dee431
ms.date: 06/08/2017
---


# Application.WindowPrev Method (Project)

Activates the window that was previously opened.


## Syntax

 _expression_. **WindowPrev**( ** _NoWrap_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _NoWrap_|Optional|**Boolean**|**True** if using **WindowPrev** on the first opened window doesn't wrap around to the last opened window. The default value is **False**.|

### Return Value

 **Boolean**


## Remarks

The window order is the order in which windows are opened. The drop-down window list in the  **Window** group of the **View** tab in the Ribbon contains the alphabetically sorted list of open windows.


