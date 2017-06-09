---
title: Application.BaselineSave Method (Project)
keywords: vbapj.chm610
f1_keywords:
- vbapj.chm610
ms.prod: project-server
api_name:
- Project.Application.BaselineSave
ms.assetid: b64967fe-f029-fc32-762a-f81cac405447
ms.date: 06/08/2017
---


# Application.BaselineSave Method (Project)

Creates a baseline plan.


## Syntax

 _expression_. **BaselineSave**( ** _All_**, ** _Copy_**, ** _Into_**, ** _RollupToSummaryTasks_**, ** _RollupFromSubtasks_**, ** _SetDefaults_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _All_|Optional|**Boolean**|**True** if the baseline plan is set for all tasks. **False** if the baseline plan is set only for the selected tasks. The default value is **True**.|
| _Copy_|Optional|**Long**| The fields to copy. Can be one of the[PjSaveBaselineFrom](pjsavebaselinefrom-enumeration-project.md) constants.|
| _Into_|Optional|**Long**|Where the fields should be copied. Can be one of the [PjSaveBaselineTo](pjsavebaselineto-enumeration-project.md) constants.|
| _RollupToSummaryTasks_|Optional|**Boolean**|**True** if parent summary task baseline data are rolled up from selected summary tasks.|
| _RollupFromSubtasks_|Optional|**Boolean**|**True** if summary task baseline data are rolled up from subtasks.|
| _SetDefaults_|Optional|**Boolean**|**True** if the values of **RollupToSummaryTasks** or **RollupFromSubtasks** are used as default values for new projects.|

### Return Value

 **Boolean**


## Remarks

 **RollupToSummaryTasks** and **RollupFromSubTasks** will have an effect only if **All** is false.


## Example

The following example first saves the baseline and then clears it.


```vb
Sub Baseline_Save() 
 
 Dim Result As Boolean 
 
 'Save baseline 
 Result = BaselineSave(True) 
 'Clear baseline 
 Result = BaselineClear (True) 
End Sub
```


