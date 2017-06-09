---
title: Cell.FieldID Property (Project)
keywords: vbapj.chm132205
f1_keywords:
- vbapj.chm132205
ms.prod: project-server
api_name:
- Project.Cell.FieldID
ms.assetid: fe7d7a7a-ebc8-4423-31de-48977cc248e1
ms.date: 06/08/2017
---


# Cell.FieldID Property (Project)

Gets the identification number of the task or resource field in the active cell. Read-only  **Long**.


## Syntax

 _expression_. **FieldID**

 _expression_ A variable that represents a **Cell** object.


## Remarks

If the active cell contains a task, can be one of the following  **[PjField](pjfield-enumeration-project.md)** constants:


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
If the active cell contains a resource, can be one of the following  **PjField** constants:


|||
|:-----|:-----|
|**pjResourceAccrueAt**|**pjResourceEnterpriseUniqueID**|
|**pjResourceActualCost**|**pjResourceFinish**|
|**pjResourceActualOvertimeCost**|**pjResourceFinish1-10**|
|**pjResourceActualOvertimeWork**|**pjResourceFlag1-20**|
|**pjResourceActualOvertimeWorkProtected**|**pjResourceGroup**|
|**pjResourceActualWork**|**pjResourceGroupBySummary**|
|**pjResourceActualWorkProtected**|**pjResourceHyperlink**|
|**pjResourceACWP**|**pjResourceHyperlinkAddress**|
|**pjResourceAssignmentDelay**|**pjResourceHyperlinkHref**|
|**pjResourceAssignmentUnits**|**pjResourceHyperlinkScreenTip**|
|**pjResourceAvailableFrom**|**pjResourceHyperlinkSubAddress**|
|**pjResourceAvailableTo**|**pjResourceID**|
|**pjResourceBaseCalendar**|**pjResourceIndex**|
|**pjResourceBaseline1-10Cost**|**pjResourceIndicators**|
|**pjResourceBaseline1-10Finish**|**pjResourceInitials**|
|**pjResourceBaseline1-10Start**|**pjResourceIsAssignment**|
|**pjResourceBaseline1-10Work**|**pjResourceLevelingDelay**|
|**pjResourceBaselineCost**|**pjResourceLinkedFields**|
|**pjResourceBaselineFinish**|**pjResourceMaterialLabel**|
|**pjResourceBaselineStart**|**pjResourceMaxUnits**|
|**pjResourceBaselineWork**|**pjResourceName**|
|**pjResourceBCWP**|**pjResourceNotes**|
|**pjResourceBCWS**|**pjResourceNumber1-20**|
|**pjResourceBookingType**|**pjResourceObjects**|
|**pjResourceCanLevel**|**pjResourceOutlineCode1-10**|
|**pjResourceCode**|**pjResourceOverallocated**|
|**pjResourceConfirmed**|**pjResourceOvertimeCost**|
|**pjResourceCost**|**pjResourceOvertimeRate**|
|**pjResourceCost1-10**|**pjResourceOvertimeWork**|
|**pjResourceCostPerUse**|**pjResourcePeakUnits**|
|**pjResourceCostRateTable**|**pjResourcePercentWorkComplete**|
|**pjResourceCostVariance**|**pjResourcePhonetics**|
|**pjResourceCreated**|**pjResourceProject**|
|**pjResourceCV**|**pjResourceRegularWork**|
|**pjResourceDate1-10**|**pjResourceRemainingCost**|
|**pjResourceDemandedRequested**|**pjResourceRemainingOvertimeCost**|
|**pjResourceDuration1-10**|**pjResourceRemainingOvertimeWork**|
|**pjResourceEMailAddress**|**pjResourceRemainingWork**|
|**pjResourceEnterprise**|**pjResourceResponsePending**|
|**pjResourceEnterpriseBaseCalendar**|**pjResourceSheetNotes**|
|**pjResourceEnterpriseCheckedOutBy**|**pjResourceStandardRate**|
|**pjResourceEnterpriseCost1-10**|**pjResourceStart**|
|**pjResourceEnterpriseDate1-30**|**pjResourceStart1-10**|
|**pjResourceEnterpriseDuration1-10**|**pjResourceSV**|
|**pjResourceEnterpriseFlag1-20**|**pjResourceTaskSummaryName**|
|**pjResourceEnterpriseGeneric**|**pjResourceTeamStatusPending**|
|**pjResourceEnterpriseInactive**|**pjResourceText1-30**|
|**pjResourceEnterpriseIsCheckedOut**|**pjResourceType**|
|**pjResourceEnterpriseLastModifiedDate**|**pjResourceUniqueID**|
|**pjResourceEnterpriseMultiValue20-29**|**pjResourceUpdateNeeded**|
|**pjResourceEnterpriseNameUsed**|**pjResourceVAC**|
|**pjResourceEnterpriseNumber1-40**|**pjResourceWindowsUserAccount**|
|**pjResourceEnterpriseOutlineCode1-29**|**pjResourceWork**|
|**pjResourceEnterpriseRBS**|**pjResourceWorkContour**|
|**pjResourceEnterpriseRequiredValues**|**pjResourceWorkgroup**|
|**pjResourceEnterpriseTeamMember**|**pjResourceWorkVariance**|
|**pjResourceEnterpriseText1-40**||

