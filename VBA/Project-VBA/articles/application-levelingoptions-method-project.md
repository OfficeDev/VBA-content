---
title: Application.LevelingOptions Method (Project)
keywords: vbapj.chm608
f1_keywords:
- vbapj.chm608
ms.prod: project-server
api_name:
- Project.Application.LevelingOptions
ms.assetid: 388a2315-e44b-3890-a16a-92ea5a778bbd
ms.date: 06/08/2017
---


# Application.LevelingOptions Method (Project)

Specifies leveling options for the active project.


## Syntax

 _expression_. **LevelingOptions**( ** _Automatic_**, ** _DelayInSlack_**, ** _AutoClearLeveling_**, ** _Order_**, ** _LevelEntireProject_**, ** _FromDate_**, ** _ToDate_**, ** _PeriodBasis_**, ** _LevelIndividualAssignments_**, ** _LevelingCanSplit_**, ** _LevelProposedBookings_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Automatic_|Optional|**Boolean**|**True** if Project automatically levels tasks in the active project.|
| _DelayInSlack_|Optional|**Boolean**|**True** if the active project can be leveled only within the available slack time. **False** if the project can be delayed in order to level resources.|
| _AutoClearLeveling_|Optional|**Boolean**|**True** if Project clears old leveling values before leveling.|
| _Order_|Optional|**Long**|A constant that specifies how Project should resolve resource conflicts when leveling tasks in the active project. Can be one of the  **[PjLevelOrder](pjlevelorder-enumeration-project.md)** constants.|
| _LevelEntireProject_|Optional|**Boolean**|**True** if the entire project is leveled. **False** if only the resources in the date range specified with FromDate and ToDate are leveled.|
| _FromDate_|Optional|**Variant**|The starting date of a range within which overallocated resources are leveled. The FromDate argument is ignored if LevelEntireProject is  **True**.|
| _ToDate_|Optional|**Variant**|The ending date of a range within which overallocated resources are leveled. The ToDate argument is ignored if LevelEntireProject is  **True**.|
| _PeriodBasis_|Optional|**Long**|Specifies how often Project should look for overallocated resources. Can be one of the  **[PjLevelPeriodBasis](pjlevelperiodbasis-enumeration-project.md)** constants.|
| _LevelIndividualAssignments_|Optional|**Boolean**|**True** if leveling can adjust individual assignments on a task.|
| _LevelingCanSplit_|Optional|**Boolean**|**True** if leveling can create splits in remaining work.|
| _LevelProposedBookings_|Optional|**Boolean**|**True** if leveling includes proposed resource bookings.|

### Return Value

 **Boolean**


## Remarks

If an argument is omitted, its default value is specified by the current setting in the  **Resource Leveling** dialog box.

Using the  **LevelingOptions** method without specifying any arguments displays the **Resource Leveling** dialog box.

To include manually scheduled tasks in the leveling options, use the  **[LevelingOptionsEx](application-levelingoptionsex-method-project.md)** method.


## Example

The following example levels resources in the application using priority to resolve conflicts.


```vb
Sub LevelOverallocatedResources() 
 LevelingOptions Order:=pjLevelPriority 
 LevelNow (True) 
End Sub
```


