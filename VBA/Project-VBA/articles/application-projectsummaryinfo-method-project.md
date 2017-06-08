---
title: Application.ProjectSummaryInfo Method (Project)
keywords: vbapj.chm601
f1_keywords:
- vbapj.chm601
ms.prod: project-server
api_name:
- Project.Application.ProjectSummaryInfo
ms.assetid: 7275598c-02b1-7e07-ecdb-04fa0a21f41a
ms.date: 06/08/2017
---


# Application.ProjectSummaryInfo Method (Project)

Sets information about a project.


## Syntax

 _expression_. **ProjectSummaryInfo**( ** _Project_**, ** _Title_**, ** _Subject_**, ** _Author_**, ** _Company_**, ** _Manager_**, ** _Keywords_**, ** _Comments_**, ** _Start_**, ** _Finish_**, ** _ScheduleFrom_**, ** _CurrentDate_**, ** _Calendar_**, ** _StatusDate_**, ** _Priority_**, ** _PartiallyDisabled_** )

 _expression_ A variable that represents an **Application** object.


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
| _Start_|Optional|**Variant**|The start date of the project. If ScheduleFrom is  **pjProjectFinish**, Start is ignored.|
| _Finish_|Optional|**Variant**|The finish date of the project. If ScheduleFrom is  **pjProjectStart**, Finish is ignored.|
| _ScheduleFrom_|Optional|**Integer**|Can be one of the following  **[PjScheduleProjectFrom](pjscheduleprojectfrom-enumeration-project.md)** constants: **pjProjectStart** or **pjProjectFinish**.|
| _CurrentDate_|Optional|**Variant**|The current date for the project.|
| _Calendar_|Optional|**String**|The name of the base calendar for the project.|
| _StatusDate_|Optional|**Variant**|The current status date for the project.|
| _Priority_|Optional|**Integer**|The priority, ranging from 0 to 1000, of the active project.|
| _PartiallyDisabled_|Optional|**Boolean**|**True** if Project displays the **Project Information** dialog box with all elements disabled except for the **Enterprise Custom Fields** section.|

### Return Value

 **Boolean**


## Remarks

Using the  **ProjectSummaryInfo** method with no arguments displays the **Project Information** dialog box.


