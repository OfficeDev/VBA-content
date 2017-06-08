---
title: Application.BaseCalendars Method (Project)
keywords: vbapj.chm604
f1_keywords:
- vbapj.chm604
ms.prod: project-server
api_name:
- Project.Application.BaseCalendars
ms.assetid: 5ae675d2-1be3-eb98-6c35-ff36c3fccf30
ms.date: 06/08/2017
---


# Application.BaseCalendars Method (Project)

Displays the  **Change Working Time** dialog box, which prompts the user to change calendar properties.


## Syntax

 _expression_. **BaseCalendars**( ** _Index_**, ** _Locked_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Optional|**String**|The calendar index number or calendar name.|
| _Locked_|Optional|**Boolean**|**True** if Project disables the **New** and **Options** buttons in the **Change Working Time** dialog box.|

### Return Value

 **Boolean**


## Remarks

The  **BaseCalendars** method has the same effect as the **Change Working Time** command on the **PROJECT** tab of the ribbon.


