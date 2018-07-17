---
title: PjTaskWarnings Enumeration (Project)
ms.prod: project-server
api_name:
- Project.PjTaskWarnings
ms.assetid: 02bff43f-4459-3c34-5e8f-c441ffefe954
ms.date: 06/08/2017
---


# PjTaskWarnings Enumeration (Project)

Contains constants that specify warnings for tasks or assignments.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
|**pjTaskWarningAssnOverallocatedInNonWorkingTime**|16384|The assignment is in overallocated non-working time.|
|**pjTaskWarningResourceBeyondMaxUnit**|64|The assignment is more than the maximum resource units available.|
|**pjTaskWarningResourceOverallocated**|128|The resource is overallocated.|
|**pjTaskWarningShadowDateDifferent**|1024|The shadow task has a different date.|
|**pjTaskWarningShadowFinishesEarlierDueToLink**|2|The shadow task finishes earlier because of a predecessor link.|
|**pjTaskWarningShadowFinishesLaterDueToLink**|1|The shadow task finishes later because of a predecessor link.|
|**pjTaskWarningShadowIncorrectByConstraintOnly**|256|The shadow task is incorrect because of a constraint.|
|**pjTaskWarningShadowIncorrectByLevelingDelayOnly**|512|The shadow task is incorrect because of a leveling delay.|
|**pjTaskWarningSubTaskFinishingAfterParentFinish**|16|The subtask finishes after the parent task.|
|**pjTaskWarningSubTaskStartingAfterParentStart**|8|The subtask starts after the parent task starts.|
|**pjTaskWarningSubTaskStartingBeforeParentStart**|4|The subtask starts before the parent task.|
|**pjTaskWarningSummaryInconsistentFinish**|2048|The finish date of the summary task is inconsistent.|
|**pjTaskWarningSummaryInconsistentStart**|32|The start date of the summary task is inconsistent.|
|**pjTaskWarningTaskFinishingInNonWorkingTime**|8192|The finish date of the task is in non-working time.|
|**pjTaskWarningTaskStartingInNonWorkingTime**|4096|The start date of the task is in non-working time.|

