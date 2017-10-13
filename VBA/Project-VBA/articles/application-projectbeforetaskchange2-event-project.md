---
title: Application.ProjectBeforeTaskChange2 Event (Project)
ms.prod: project-server
api_name:
- Project.Application.ProjectBeforeTaskChange2
ms.assetid: 00992e39-dcbd-3826-4ce6-e2be55dc9c2c
ms.date: 06/08/2017
---


# Application.ProjectBeforeTaskChange2 Event (Project)

Occurs before the user changes the value of a task field. Uses the **EventInfo** object parameter.


## Syntax

_expression_. **ProjectBeforeTaskChange2** (**_tsk_**, **_Field_**, **_NewVal_**, **_Info_**)

_expression_ A variable that represents an **Application** object.


### Parameters

|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _tsk_|Required|**Task**|The task whose field is being changed.|
| _Field_|Required|**PjField**|The field being changed. If more than one field is changed by the user, the event is fired for each field changed. Can be one of the **[PjField constants](#pjfield-constants)**.|
| _NewVal_|Required|**Variant**|The new value for the field specified with _Field_.|
| _Info_|Required|**EventInfo**|EventInfo.Cancel is **False** when the event occurs. If the event procedure sets this argument to **True**, the value for the field specified with _Field_ is not changed.|

<br/>

#### PjField constants

|||
|:-----|:-----|
|**pjTaskActualCost**|**pjTaskHyperlinkSubAddress**|
|**pjTaskActualDuration**|**pjTaskID**|
|**pjTaskActualFinish**|**pjTaskIgnoreResourceCalendar**|
|**pjTaskActualOvertimeCost**|**pjTaskIndex**|
|**pjTaskActualOvertimeWork**|**pjTaskIndicators**|
|**pjTaskActualOvertimeWorkProtected**|**pjTaskIsAssignment**|
|**pjTaskActualStart**|**pjTaskLateFinish**|
|**pjTaskActualWork**|**pjTaskLateStart**|
|**pjTaskActualWorkProtected**|**pjTaskLevelAssignments**|
|**pjTaskACWP**|**pjTaskLevelCanSplit**|
|**pjTaskAssignmentDelay**|**pjTaskLevelDelay**|
|**pjTaskAssignmentUnits**|**pjTaskLinkedFields**|
|**pjTaskBaseline1-10Cost**|**pjTaskMarked**|
|**pjTaskBaseline1-10Duration**|**pjTaskMilestone**|
|**pjTaskBaseline1-10DurationEstimated**|**pjTaskName**|
|**pjTaskBaseline1-10Finish**|**pjTaskNotes**|
|**pjTaskBaseline1-10Start**|**pjTaskNumber1-20**|
|**pjTaskBaseline1-10Work**|**pjTaskObjects**|
|**pjTaskBaselineCost**|**pjTaskOutlineCode1-10**|
|**pjTaskBaselineDuration**|**pjTaskOutlineLevel**|
|**pjTaskBaselineDurationEstimated**|**pjTaskOutlineNumber**|
|**pjTaskBaselineFinish**|**pjTaskOverallocated**|
|**pjTaskBaselineStart**|**pjTaskOvertimeCost**|
|**pjTaskBaselineWork**|**pjTaskOvertimeWork**|
|**pjTaskBCWP**|**pjTaskParentTask**|
|**pjTaskBCWS**|**pjTaskPercentComplete**|
|**pjTaskCalendar**|**pjTaskPercentWorkComplete**|
|**pjTaskConfirmed**|**pjTaskPhysicalPercentComplete**|
|**pjTaskConstraintDate**|**pjTaskPredecessors**|
|**pjTaskConstraintType**|**pjTaskPreleveledFinish**|
|**pjTaskContact**|**pjTaskPreleveledStart**|
|**pjTaskCost**|**pjTaskPriority**|
|**pjTaskCost1-10**|**pjTaskProject**|
|**pjTaskCostRateTable**|**pjTaskRecurring**|
|**pjTaskCostVariance**|**pjTaskRegularWork**|
|**pjTaskCPI**|**pjTaskRemainingCost**|
|**pjTaskCreated**|**pjTaskRemainingDuration**|
|**pjTaskCritical**|**pjTaskRemainingOvertimeCost**|
|**pjTaskCV**|**pjTaskRemainingOvertimeWork**|
|**pjTaskCVPercent**|**pjTaskRemainingWork**|
|**pjTaskDate1-10**|**pjTaskResourceEnterpriseMultiValueCode20-29**|
|**pjTaskDeadline**|**pjTaskResourceEnterpriseOutlineCode1-29**|
|**pjTaskDelay**|**pjTaskResourceEnterpriseRBS**|
|**pjTaskDemandedRequest**|**pjTaskResourceGroup**|
|**pjTaskDuration**|**pjTaskResourceInitials**|
|**pjTaskDuration1-10**|**pjTaskResourceNames**|
|**pjTaskDuration1-10Estimated**|**pjTaskResourcePhonetics**|
|**pjTaskDurationVariance**|**pjTaskResourceType**|
|**pjTaskEAC**|**pjTaskResponsePending**|
|**pjTaskEarlyFinish**|**pjTaskResume**|
|**pjTaskEarlyStart**|**pjTaskResumeNoEarlierThan**|
|**pjTaskEarnedValueMethod**|**pjTaskRollup**|
|**pjTaskEffortDriven**|**pjTaskSheetNotes**|
|**pjTaskEnterpriseCost1-10**|**pjTaskSPI**|
|**pjTaskEnterpriseDate1-30**|**pjTaskStart**|
|**pjTaskEnterpriseDuration1-10**|**pjTaskStart1-10**|
|**pjTaskEnterpriseFlag1-20**|**pjTaskStartSlack**|
|**pjTaskEnterpriseNumber1-40**|**pjTaskStartVariance**|
|**pjTaskEnterpriseOutlineCode1-30**|**pjTaskStatus**|
|**pjTaskEnterpriseProjectCost1-10**|**pjTaskStatusIndicator**|
|**pjTaskEnterpriseProjectDate1-30**|**pjTaskStop**|
|**pjTaskEnterpriseProjectDuration1-10**|**pjTaskSubproject**|
|**pjTaskEnterpriseProjectFlag1-20**|**pjTaskSubprojectReadOnly**|
|**pjTaskEnterpriseProjectNumber1-40**|**pjTaskSuccessors**|
|**pjTaskEnterpriseProjectOutlineCode1-30**|**pjTaskSummary**|
|**pjTaskEnterpriseProjectText1-40**|**pjTaskSV**|
|**pjTaskEnterpriseText1-40**|**pjTaskSVPercent**|
|**pjTaskEstimated**|**pjTaskTCPI**|
|**pjTaskExternalTask**|**pjTaskTeamStatusPending**|
|**pjTaskFinish**|**pjTaskText1-30**|
|**pjTaskFinish1-10**|**pjTaskTotalSlack**|
|**pjTaskFinishSlack**|**pjTaskType**|
|**pjTaskFinishVariance**|**pjTaskUniqueID**|
|**pjTaskFixedCost**|**pjTaskUniquePredecessors**|
|**pjTaskFixedCostAccrual**|**pjTaskUniqueSuccessors**|
|**pjTaskFixedDuration**|**pjTaskUpdateNeeded**|
|**pjTaskFlag1-20**|**pjTaskVAC**|
|**pjTaskFreeSlack**|**pjTaskWBS**|
|**pjTaskGroupBySummary**|**pjTaskWBSPredecessors**|
|**pjTaskHideBar**|**pjTaskWBSSuccessors**|
|**pjTaskHyperlink**|**pjTaskWork**|
|**pjTaskHyperlinkAddress**|**pjTaskWorkContour**|
|**pjTaskHyperlinkHref**|**pjTaskWorkVariance**|
|**pjTaskHyperlinkScreenTip**||

<br/>

### Return value

Nothing


## Remarks

Project events do not occur when the project is embedded in another document or application. For more information and sample code for creating and testing an event handler, see [Using Events with Application and Project Objects](using-events-with-application-and-project-objects.md).

The **ProjectBeforeTaskChange2** event doesn't occur when timescaled data changes, when constraint data in the Task Details Form changes, when a task is split by manipulating its task bar on the Gantt Chart, when changes are made to outline level or outline number, when a baseline is saved, when a baseline is cleared, when an entire task row is pasted, during resource pool operations, when inserting or removing a subproject, or when changes have been made by using a custom form.


