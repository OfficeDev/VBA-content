---
title: Application.ProjectBeforeSaveBaseline Event (Project)
ms.prod: project-server
api_name:
- Project.Application.ProjectBeforeSaveBaseline
ms.assetid: bcdd2134-03dd-e26d-66db-095bda6a7162
ms.date: 06/08/2017
---


# Application.ProjectBeforeSaveBaseline Event (Project)

Occurs before a baseline is saved. Uses the  **EventInfo** object parameter.


## Syntax

 _expression_. **ProjectBeforeSaveBaseline**( ** _pj_**, ** _Interim_**, ** _bl_**, ** _InterimCopy_**, ** _InterimInto_**, ** _AllTasks_**, ** _RollupToSummaryTasks_**, ** _RollupFromSubtasks_**, ** _Info_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _pj_|Required|**Project**|The project displayed in the window.|
| _Interim_|Required|**Boolean**|**True** if saving an interim baseline. **False** if saving a full baseline.|
| _bl_|Required|**PjBaselines**|The baseline you are saving. Can be one of the following  **PjBaselines** constants: **pjBaseline**, **pjBaseline1**, **pjBaseline2**, **pjBaseline3**, **pjBaseline4**, **pjBaseline5**, **pjBaseline6**, **pjBaseline7**, **pjBaseline8**, **pjBaseline9**, or **pjBaseline10**.|
| _InterimCopy_|Required|**PjSaveBaselineFrom**|The interim plan being copied from. Can be one of the following  **PjSaveBaselineFrom** constants: **pjCopyBaseline**, **pjCopyBaseline1**, **pjCopyBaseline2**, **pjCopyBaseline3**, **pjCopyBaseline4**, **pjCopyBaseline5**, **pjCopyBaseline6**, **pjCopyBaseline7**, **pjCopyBaseline8**, **pjCopyBaseline9**, **pjCopyBaseline10**, **pjCopyCurrent**, **pjCopyStart_Finish1**, **pjCopyStart_Finish2**, **pjCopyStart_Finish3**, ** pjCopyStart_Finish4**, **pjCopyStart_Finish5**, **pjCopyStart_Finish6**, **pjCopyStart_Finish7**, **pjCopyStart_Finish8**, **pjCopyStart_Finish9**, or **pjCopyStart_Finish10**.|
| _InterimInto_|Required|**PjSaveBaselineTo**|The interim plan to which you are saving. Can be one of the following  **PjSaveBaselineTo** constants: **pjIntoBaseline**, **pjIntoBaseline1**, **pjIntoBaseline2**, **pjIntoBaseline3**, **pjIntoBaseline4**, **pjIntoBaseline5**, **pjIntoBaseline6**, **pjIntoBaseline7**, **pjIntoBaseline8**, **pjIntoBaseline9**, **pjIntoBaseline10**, ** pjIntoStart_Finish1**, **pjIntoStart_Finish2**, **pjIntoStart_Finish3**, **pjIntoStart_Finish4**, **pjIntoStart_Finish5**, **pjIntoStart_Finish6**, **pjIntoStart_Finish7**, **pjIntoStart_Finish8**, **pjIntoStart_Finish9**, or **pjIntoStart_Finish10**.|
| _AllTasks_|Required|**Boolean**|**True** if saving the entire project.|
| _RollupToSummaryTasks_|Required|**Boolean**|**True** if you wish to roll up baselines to summary tasks.|
| _RollupFromSubtasks_|Required|**Boolean**|**True** if you wish to roll up baselines from subtasks.|
| _Info_|Required|**EventInfo**|EventInfo.Cancel is  **False** when the event occurs. If the event procedure sets this argument to **True**, the baseline is not saved.|

### Return Value

nothing


## Remarks

Project events do not occur when the project is embedded in another document or application.


