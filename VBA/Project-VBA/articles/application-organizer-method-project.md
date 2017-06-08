---
title: Application.Organizer Method (Project)
keywords: vbapj.chm126
f1_keywords:
- vbapj.chm126
ms.prod: project-server
api_name:
- Project.Application.Organizer
ms.assetid: 4269290c-7be9-a0af-526d-bde73114c24b
ms.date: 06/08/2017
---


# Application.Organizer Method (Project)

Displays the  **Organizer** dialog box, which enables the user to manage views, reports, modules, tables, filters, calendars, maps, fields, and groups.


## Syntax

 _expression_. **Organizer**( ** _Type_**, ** _Task_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Optional|**Long**|The type of item to manage. Can be one of the  **[PjOrganizer](pjorganizer-enumeration-project.md)** constants. The default value is **pjViews**.|
| _Task_|Optional|**Boolean**|**True** if the item applies to tasks. **False** if the item applies to resources. The default value is **True**.|

### Return Value

 **Boolean**


## Remarks

If  _Type_ is set to **pjToolbar**, that maps to the **Modules** tab in the **Organizer** dialog box.


