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


|                                                         |                                                              |
|:--------------------------------------------------------|:-------------------------------------------------------------|
| <strong>pjTaskActualCost</strong>                       | <strong>pjTaskHyperlinkSubAddress</strong>                   |
| <strong>pjTaskActualDuration</strong>                   | <strong>pjTaskID</strong>                                    |
| <strong>pjTaskActualFinish</strong>                     | <strong>pjTaskIgnoreResourceCalendar</strong>                |
| <strong>pjTaskActualOvertimeCost</strong>               | <strong>pjTaskIndex</strong>                                 |
| <strong>pjTaskActualOvertimeWork</strong>               | <strong>pjTaskIndicators</strong>                            |
| <strong>pjTaskActualOvertimeWorkProtected</strong>      | <strong>pjTaskIsAssignment</strong>                          |
| <strong>pjTaskActualStart</strong>                      | <strong>pjTaskLateFinish</strong>                            |
| <strong>pjTaskActualWork</strong>                       | <strong>pjTaskLateStart</strong>                             |
| <strong>pjTaskActualWorkProtected</strong>              | <strong>pjTaskLevelAssignments</strong>                      |
| <strong>pjTaskACWP</strong>                             | <strong>pjTaskLevelCanSplit</strong>                         |
| <strong>pjTaskAssignmentDelay</strong>                  | <strong>pjTaskLevelDelay</strong>                            |
| <strong>pjTaskAssignmentUnits</strong>                  | <strong>pjTaskLinkedFields</strong>                          |
| <strong>pjTaskBaseline1-10Cost</strong>                 | <strong>pjTaskMarked</strong>                                |
| <strong>pjTaskBaseline1-10Duration</strong>             | <strong>pjTaskMilestone</strong>                             |
| <strong>pjTaskBaseline1-10DurationEstimated</strong>    | <strong>pjTaskName</strong>                                  |
| <strong>pjTaskBaseline1-10Finish</strong>               | <strong>pjTaskNotes</strong>                                 |
| <strong>pjTaskBaseline1-10Start</strong>                | <strong>pjTaskNumber1-20</strong>                            |
| <strong>pjTaskBaseline1-10Work</strong>                 | <strong>pjTaskObjects</strong>                               |
| <strong>pjTaskBaselineCost</strong>                     | <strong>pjTaskOutlineCode1-10</strong>                       |
| <strong>pjTaskBaselineDuration</strong>                 | <strong>pjTaskOutlineLevel</strong>                          |
| <strong>pjTaskBaselineDurationEstimated</strong>        | <strong>pjTaskOutlineNumber</strong>                         |
| <strong>pjTaskBaselineFinish</strong>                   | <strong>pjTaskOverallocated</strong>                         |
| <strong>pjTaskBaselineStart</strong>                    | <strong>pjTaskOvertimeCost</strong>                          |
| <strong>pjTaskBaselineWork</strong>                     | <strong>pjTaskOvertimeWork</strong>                          |
| <strong>pjTaskBCWP</strong>                             | <strong>pjTaskParentTask</strong>                            |
| <strong>pjTaskBCWS</strong>                             | <strong>pjTaskPercentComplete</strong>                       |
| <strong>pjTaskCalendar</strong>                         | <strong>pjTaskPercentWorkComplete</strong>                   |
| <strong>pjTaskConfirmed</strong>                        | <strong>pjTaskPhysicalPercentComplete</strong>               |
| <strong>pjTaskConstraintDate</strong>                   | <strong>pjTaskPredecessors</strong>                          |
| <strong>pjTaskConstraintType</strong>                   | <strong>pjTaskPreleveledFinish</strong>                      |
| <strong>pjTaskContact</strong>                          | <strong>pjTaskPreleveledStart</strong>                       |
| <strong>pjTaskCost</strong>                             | <strong>pjTaskPriority</strong>                              |
| <strong>pjTaskCost1-10</strong>                         | <strong>pjTaskProject</strong>                               |
| <strong>pjTaskCostRateTable</strong>                    | <strong>pjTaskRecurring</strong>                             |
| <strong>pjTaskCostVariance</strong>                     | <strong>pjTaskRegularWork</strong>                           |
| <strong>pjTaskCPI</strong>                              | <strong>pjTaskRemainingCost</strong>                         |
| <strong>pjTaskCreated</strong>                          | <strong>pjTaskRemainingDuration</strong>                     |
| <strong>pjTaskCritical</strong>                         | <strong>pjTaskRemainingOvertimeCost</strong>                 |
| <strong>pjTaskCV</strong>                               | <strong>pjTaskRemainingOvertimeWork</strong>                 |
| <strong>pjTaskCVPercent</strong>                        | <strong>pjTaskRemainingWork</strong>                         |
| <strong>pjTaskDate1-10</strong>                         | <strong>pjTaskResourceEnterpriseMultiValueCode20-29</strong> |
| <strong>pjTaskDeadline</strong>                         | <strong>pjTaskResourceEnterpriseOutlineCode1-29</strong>     |
| <strong>pjTaskDelay</strong>                            | <strong>pjTaskResourceEnterpriseRBS</strong>                 |
| <strong>pjTaskDemandedRequest</strong>                  | <strong>pjTaskResourceGroup</strong>                         |
| <strong>pjTaskDuration</strong>                         | <strong>pjTaskResourceInitials</strong>                      |
| <strong>pjTaskDuration1-10</strong>                     | <strong>pjTaskResourceNames</strong>                         |
| <strong>pjTaskDuration1-10Estimated</strong>            | <strong>pjTaskResourcePhonetics</strong>                     |
| <strong>pjTaskDurationVariance</strong>                 | <strong>pjTaskResourceType</strong>                          |
| <strong>pjTaskEAC</strong>                              | <strong>pjTaskResponsePending</strong>                       |
| <strong>pjTaskEarlyFinish</strong>                      | <strong>pjTaskResume</strong>                                |
| <strong>pjTaskEarlyStart</strong>                       | <strong>pjTaskResumeNoEarlierThan</strong>                   |
| <strong>pjTaskEarnedValueMethod</strong>                | <strong>pjTaskRollup</strong>                                |
| <strong>pjTaskEffortDriven</strong>                     | <strong>pjTaskSheetNotes</strong>                            |
| <strong>pjTaskEnterpriseCost1-10</strong>               | <strong>pjTaskSPI</strong>                                   |
| <strong>pjTaskEnterpriseDate1-30</strong>               | <strong>pjTaskStart</strong>                                 |
| <strong>pjTaskEnterpriseDuration1-10</strong>           | <strong>pjTaskStart1-10</strong>                             |
| <strong>pjTaskEnterpriseFlag1-20</strong>               | <strong>pjTaskStartSlack</strong>                            |
| <strong>pjTaskEnterpriseNumber1-40</strong>             | <strong>pjTaskStartVariance</strong>                         |
| <strong>pjTaskEnterpriseOutlineCode1-30</strong>        | <strong>pjTaskStatus</strong>                                |
| <strong>pjTaskEnterpriseProjectCost1-10</strong>        | <strong>pjTaskStatusIndicator</strong>                       |
| <strong>pjTaskEnterpriseProjectDate1-30</strong>        | <strong>pjTaskStop</strong>                                  |
| <strong>pjTaskEnterpriseProjectDuration1-10</strong>    | <strong>pjTaskSubproject</strong>                            |
| <strong>pjTaskEnterpriseProjectFlag1-20</strong>        | <strong>pjTaskSubprojectReadOnly</strong>                    |
| <strong>pjTaskEnterpriseProjectNumber1-40</strong>      | <strong>pjTaskSuccessors</strong>                            |
| <strong>pjTaskEnterpriseProjectOutlineCode1-30</strong> | <strong>pjTaskSummary</strong>                               |
| <strong>pjTaskEnterpriseProjectText1-40</strong>        | <strong>pjTaskSV</strong>                                    |
| <strong>pjTaskEnterpriseText1-40</strong>               | <strong>pjTaskSVPercent</strong>                             |
| <strong>pjTaskEstimated</strong>                        | <strong>pjTaskTCPI</strong>                                  |
| <strong>pjTaskExternalTask</strong>                     | <strong>pjTaskTeamStatusPending</strong>                     |
| <strong>pjTaskFinish</strong>                           | <strong>pjTaskText1-30</strong>                              |
| <strong>pjTaskFinish1-10</strong>                       | <strong>pjTaskTotalSlack</strong>                            |
| <strong>pjTaskFinishSlack</strong>                      | <strong>pjTaskType</strong>                                  |
| <strong>pjTaskFinishVariance</strong>                   | <strong>pjTaskUniqueID</strong>                              |
| <strong>pjTaskFixedCost</strong>                        | <strong>pjTaskUniquePredecessors</strong>                    |
| <strong>pjTaskFixedCostAccrual</strong>                 | <strong>pjTaskUniqueSuccessors</strong>                      |
| <strong>pjTaskFixedDuration</strong>                    | <strong>pjTaskUpdateNeeded</strong>                          |
| <strong>pjTaskFlag1-20</strong>                         | <strong>pjTaskVAC</strong>                                   |
| <strong>pjTaskFreeSlack</strong>                        | <strong>pjTaskWBS</strong>                                   |
| <strong>pjTaskGroupBySummary</strong>                   | <strong>pjTaskWBSPredecessors</strong>                       |
| <strong>pjTaskHideBar</strong>                          | <strong>pjTaskWBSSuccessors</strong>                         |
| <strong>pjTaskHyperlink</strong>                        | <strong>pjTaskWork</strong>                                  |
| <strong>pjTaskHyperlinkAddress</strong>                 | <strong>pjTaskWorkContour</strong>                           |
| <strong>pjTaskHyperlinkHref</strong>                    | <strong>pjTaskWorkVariance</strong>                          |
| <strong>pjTaskHyperlinkScreenTip</strong>               |                                                              |

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

