---
title: Application.CalendarDateShading Method (Project)
keywords: vbapj.chm2344
f1_keywords:
- vbapj.chm2344
ms.prod: project-server
api_name:
- Project.Application.CalendarDateShading
ms.assetid: fedb04c6-e9a4-9289-aedd-042f3751e27d
ms.date: 06/08/2017
---


# Application.CalendarDateShading Method (Project)

Determines which calendar is used when determining when and how dates are shaded in the Calendar view.


## Syntax

 _expression_. **CalendarDateShading**( ** _BaseCalendarName_**, ** _ResourceUniqueID_**, ** _ProjectIndex_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _BaseCalendarName_|Optional|**String**|If referring to a single project, or the master project in a consolidated project, the name of a base calendar to use for shading. If referring to an subproject in a consolidated project, the name of a base calendar and the name of the subproject in the manner of " **Calendar** [ **Project** ]", where **Calendar** is the name of the base calendar and **Project** is the name of the subproject.|
| _ResourceUniqueID_|Optional|**Long**|The unique identification number of a resource. The corresponding resource calendar is used for shading.|
| _ProjectIndex_|Optional|**Variant**|Due to changes in the Project object model, this argument no longer has an effect. It has been retained for backward compatibility.|

### Return Value

 **Boolean**


## Remarks

When the Calendar view is active, using the  **CalendarDateShading** method with no arguments displays the **Timescale** dialog box with the **Date Shading** tab selected. You must specify either **BaseCalendarName** or **ResourceUniqueID**, but you cannot specify both.


