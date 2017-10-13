---
title: Application.ProjectBeforeResourceChange2 Event (Project)
ms.prod: project-server
api_name:
- Project.Application.ProjectBeforeResourceChange2
ms.assetid: 84128c94-0d0d-f8f2-6d5a-4c05a61a0a8d
ms.date: 06/08/2017
---


# Application.ProjectBeforeResourceChange2 Event (Project)

Occurs before the user changes the value of a resource field. Uses the **EventInfo** object parameter.


## Syntax

_expression_. **ProjectBeforeResourceChange2** (**_res_**, **_Field_**, **_NewVal_**, **_Info_**)

_expression_ A variable that represents an **Application** object.


### Parameters

|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _res_|Required|**Resource**|The resource whose field is being changed.|
| _Field_|Required|**Long**|The field being changed. If more than one field is changed by the user, the event is fired for each field changed. Can be one of the **[PjField constants](#pjfield-constants)**.|
| _NewVal_|Required|**Variant**|The new value for the field specified with Field.|
| _Info_|Required|**EventInfo**|EventInfo.Cancel is **False** when the event occurs. If the event procedure sets this argument to **True**, the value for the field specified with Field is not changed.|

<br/>

#### PjField constants

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

<br/>

### Return value

Nothing


## Remarks

Project events do not occur when the project is embedded in another document or application.

The **ProjectBeforeResourceChange2** event doesn't occur when timescaled data changes, when a baseline is cleared, when an entire resource row is pasted, during resource pool operations, when inserting or removing a subproject, or when changes have been made by using a custom form.


