---
title: Application.ProjectSummaryInfoEx Method (Project)
keywords: vbapj.chm634
f1_keywords:
- vbapj.chm634
ms.assetid: 2827f735-6a7b-9f33-c1c6-2c5f1f7492f6
ms.date: 06/08/2017
ms.prod: project-server
---


# Application.ProjectSummaryInfoEx Method (Project)

Returns information about project summary, including the Project Utilization type and Project Utilization date information. Introduced in Office 2016.


## Syntax

 _expression_. **ProjectSummaryInfoEx**( _Project_,  _Project_,  _Title_,  _Subject_,  _Author_,  _Company_,  _Manager_,  _Keywords_,  _Comments_,  _Start_,  _Finish_,  _ScheduleFrom_,  _CurrentDate_,  _Calendar_,  _StatusDate_,  _Priority_,  _UtilizationType_,  _UtilizationDate_,  _PartiallyDisabled_)

 _expression_ A variable that represents a **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Project_|Optional|**String**|The file name of the project that should have its project information edited.|
| _Title_|Optional|**String**|The title of the project.|
| _Subject_|Optional|**String**|The subject of the project.|
| _Author_|Optional|**String**|The author of the project.|
| _Company_|Optional|**String**|The company associated with the project.|
| _Manager_|Optional|**String**|The manager of the project.|
| _Keywords_|Optional|**String**|The keywords associated with the project.|
| _Comments_|Optional|**String**|The comments associated with the project.|
| _Start_|Optional|**Variant**|The start date of the project. If ScheduleFrom is pjProjectFinish, Start is ignored|
| _Finish_|Optional|**Variant**|The start date of the project. If  **ScheduleFrom** is **pjProjectFinish**,  _Start_ is ignored|
| _ScheduleFrom_|Optional|**Integer**|Can be one of the following  **PjScheduleProjectFrom** constants: **pjProjectStart** or **pjProjectFinish**.|
| _CurrentDate_|Optional|**Variant**|The current date for the project.|
| _Calendar_|Optional|**String**|The name of the base calendar for the project.|
| _StatusDate_|Optional|**Variant**|The current status date for the project.|
| _Priority_|Optional|**Integer**|The priority, ranging from 0 to 1000, of the active project.|
| _UtilizationType_|Optional|**Variant**||
| _UtilizationDate_|Optional|**Variant**||
| _PartiallyDisabled_|Optional|**Boolean**|**True** if Project displays the **Project Information** dialog box with all elements disabled except for the **Enterprise Custom Fields** section.|

### Return Value

 **BOOL**


### Remarks

Using the  **ProjectSummaryInfoEx** method with no arguments displays the **Project Information** dialog box


