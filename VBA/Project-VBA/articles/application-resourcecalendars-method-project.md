---
title: Application.ResourceCalendars Method (Project)
keywords: vbapj.chm605
f1_keywords:
- vbapj.chm605
ms.prod: project-server
api_name:
- Project.Application.ResourceCalendars
ms.assetid: 8c40cfad-ec40-43a4-5698-de5abaea7243
ms.date: 06/08/2017
---


# Application.ResourceCalendars Method (Project)

Displays the  **Change Working Time** dialog box, which prompts the user to manage calendars.


## Syntax

 _expression_. **ResourceCalendars**( ** _Index_**, ** _Locked_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Optional|**String**|The resource index number or resource name.|
| _Locked_|Optional|**Boolean**|**False** if the user can set working time for selected dates for a resource. **True** if the fields are locked for editing. The default value is **False**.|

### Return Value

 **Boolean**


## Remarks

The  **ResourceCalendars** method returns a trappable error (error code 1101) when applied to material resources.


