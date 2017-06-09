---
title: Application.CalendarLayout Method (Project)
keywords: vbapj.chm2346
f1_keywords:
- vbapj.chm2346
ms.prod: project-server
api_name:
- Project.Application.CalendarLayout
ms.assetid: c948c118-c50f-493d-ba3a-e43ee0d50fa3
ms.date: 06/08/2017
---


# Application.CalendarLayout Method (Project)

Changes how task bars are arranged on the Calendar.


## Syntax

 _expression_. **CalendarLayout**( ** _SortOrder_**, ** _AutoLayout_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SortOrder_|Optional|**Boolean**|**True** if tasks are displayed in the Calendar using the current sort order. **False** if the sort order changes to display as many tasks as possible. The default value is **True**.|
| _AutoLayout_|Optional|**Boolean**|**True** if the Calendar view automatically changes to reflect task changes.|

### Return Value

 **Boolean**


