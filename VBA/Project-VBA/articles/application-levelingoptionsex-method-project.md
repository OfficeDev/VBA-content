---
title: Application.LevelingOptionsEx Method (Project)
keywords: vbapj.chm2249
f1_keywords:
- vbapj.chm2249
ms.prod: project-server
api_name:
- Project.Application.LevelingOptionsEx
ms.assetid: f8799750-fecf-48d1-7559-25cd7a8d3d28
ms.date: 06/08/2017
---


# Application.LevelingOptionsEx Method (Project)

Specifies leveling options for the active project, including leveling of manually scheduled tasks.


## Syntax

 _expression_. **LevelingOptionsEx**( ** _Automatic_**, ** _DelayInSlack_**, ** _AutoClearLeveling_**, ** _Order_**, ** _LevelEntireProject_**, ** _FromDate_**, ** _ToDate_**, ** _PeriodBasis_**, ** _LevelIndividualAssignments_**, ** _LevelingCanSplit_**, ** _LevelProposedBookings_**, ** _LevelPinnedTasks_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Automatic_|Optional|**Boolean**|**True** if Project automatically levels tasks in the active project. **False** if leveling is manual. The default value is **False**.|
| _DelayInSlack_|Optional|**Boolean**|**True** if the active project can be leveled only within the available slack time. **False** if the project can be delayed to level resources. The default value is **False**.|
| _AutoClearLeveling_|Optional|**Boolean**|**True** if Project clears old leveling values before leveling; otherwise, **False**. The default value is **True**.|
| _Order_|Optional|**Integer**|A constant that specifies how Project should resolve resource conflicts when leveling tasks in the active project. Can be one of the  **[PjLevelOrder](pjlevelorder-enumeration-project.md)** constants. The default value is **pjLevelOrderStandard**.|
| _LevelEntireProject_|Optional|**Boolean**|**True** if the entire project is leveled. **False** if only the resources in the date range specified with _FromDate_ and _ToDate_ are leveled. The default value is **True**.|
| _FromDate_|Optional|**Variant**|The starting date of a range within which overallocated resources are leveled. The  _FromDate_ argument is ignored if _LevelEntireProject_ is **True**.|
| _ToDate_|Optional|**Variant**|The ending date of a range within which overallocated resources are leveled. The  _ToDate_ argument is ignored if _LevelEntireProject_ is **True**.|
| _PeriodBasis_|Optional|**Integer**|Specifies how often Project should look for overallocated resources. Can be one of the  **[PjLevelPeriodBasis](pjlevelperiodbasis-enumeration-project.md)** constants. The default value is **pjDayByDay**.|
| _LevelIndividualAssignments_|Optional|**Boolean**|**True** if leveling can adjust individual assignments on a task; otherwise, **false**. The default value is **True**.|
| _LevelingCanSplit_|Optional|**Boolean**|**True** if leveling can create splits in remaining work; otherwise, **False**. The default value is **True**.|
| _LevelProposedBookings_|Optional|**Boolean**|**True** if leveling includes proposed resource bookings; otherwise, **False**. The default value is **False**.|
| _LevelPinnedTasks_|Optional|**Boolean**|**True** if manually scheduled tasks are leveled; otherwise, **False**. The default value is **False**.|

### Return Value

 **Boolean**


## Remarks

If an argument is omitted, its default value is specified by the current setting in the  **Resource Leveling** dialog box. The default values specified in the list of parameters are for a new installation of Project. To access the **Resource Leveling** dialog box, on the **Resource** tab of the ribbon, choose **Leveling Options**.

Using the  **LevelingOptionsEx** method with no arguments displays the **Resource Leveling** dialog box.

To get or set only the option for leveling manually scheduled tasks, see the  **[LevelFreeformTasks](application-levelfreeformtasks-property-project.md)** property.


## Example

The following example levels only selected resources for tasks within August 2012, by using task priority to resolve conflicts.


```vb
Sub LevelOverallocatedResources() 
    LevelingOptionsEx Order:=pjLevelPriority, LevelEntireProject:=False, _ 
        FromDate:="8/1/2012", ToDate:="8/31/2012" 
    LevelNow (False) 
End Sub
```


