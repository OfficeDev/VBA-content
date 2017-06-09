---
title: Application.OptionsCalculation Method (Project)
keywords: vbapj.chm606
f1_keywords:
- vbapj.chm606
ms.prod: project-server
api_name:
- Project.Application.OptionsCalculation
ms.assetid: 608d5bd2-eb6b-0e3c-789a-c376ee55816d
ms.date: 06/08/2017
---


# Application.OptionsCalculation Method (Project)

Sets calculation options.


## Syntax

 _expression_. **OptionsCalculation**( ** _Automatic_**, ** _AutoTrack_**, ** _SpreadPercentToStatusDate_**, ** _SpreadCostsToStatusDate_**, ** _AutoCalcCosts_**, ** _FixedCostAccrual_**, ** _CalcMultipleCriticalPaths_**, ** _CriticalSlack_**, ** _SetDefaults_**, ** _CalcInsProjLikeSummTask_**, ** _MoveCompleted_**, ** _AndMoveRemaining_**, ** _MoveRemaining_**, ** _AndMoveCompleted_**, ** _EVMethod_**, ** _EVBaseline_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Automatic_|Optional|**Boolean**|**True** if the calculation mode is automatic.|
| _AutoTrack_|Optional|**Boolean**|**True** if task tracking fields automatically update resource assignments.|
| _SpreadPercentToStatusDate_|Optional|**Boolean**|**True** if edits to total task percent complete are spread to the status date.|
| _SpreadCostsToStatusDate_|Optional|**Boolean**|**True** if edits to total actual cost are spread to the status date. The SpreadCostsToStatusDate argument is ignored if AutoCalcCosts is **True**.|
| _AutoCalcCosts_|Optional|**Boolean**|**True** if actual costs are always calculated by Project.|
| _FixedCostAccrual_|Optional|**Long**|The default method used to accrue fixed costs. Can be one of the following  **[PjAccrueAt](pjaccrueat-enumeration-project.md)** constants: **pjStart**, **pjEnd**, or **pjProrated**.|
| _CalcMultipleCriticalPaths_|Optional|**Boolean**|**True** if Project calculates multiple critical paths for the project.|
| _CriticalSlack_|Optional|**Integer**|The maximum amount of slack allowed for critical tasks.|
| _SetDefaults_|Optional|**Boolean**|**True** if the values specified for all arguments except Automatic become the default for new projects.|
| _CalcInsProjLikeSummTask_|Optional|**Boolean**|**True** if subprojects in a master project behave like normal summary tasks. **False** if subprojects are calculated on their own. The default value is **False**.|
| _MoveCompleted_|Optional|**Boolean**|**True** if Project moves the end of completed parts after the status date back to the status date.|
| _AndMoveRemaining_|Optional|**Boolean**|**True** if Project moves the start of remaining parts back to the status date.|
| _MoveRemaining_|Optional|**Boolean**|**True** if Project moves the start of remaining parts before the status date forward to the status date.|
| _AndMoveCompleted_|Optional|**Boolean**|**True** if Project moves the end of completed parts forward to the status date.|
| _EVMethod_|Optional|**Long**|The default method for calculating earned value. Can be one of the following  **[PJEarnedValueMethod](pjearnedvaluemethod-enumeration-project.md)** constants: **pjPercentComplete** or **pjPhysicalPercentComplete**.|
| _EVBaseline_|Optional|**Long**|The baseline to use when calculating earned value. Can be one of the following  **[PjBaselines](pjbaselines-enumeration-project.md)** constants: **pjBaseline**, or **pjBaseline1**. . . **pjBaseline10**.|

### Return Value

 **Boolean**


## Remarks

If an argument is omitted, its default value is specified by the setting on the  **Schedule** tab of the **Project Options** dialog box.

Using the  **OptionsCalculation** method without specifying any arguments displays the **Project Options** dialog box with the **General** tab selected.




