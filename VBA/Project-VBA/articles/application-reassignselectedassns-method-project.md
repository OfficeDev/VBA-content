---
title: Application.ReassignSelectedAssns Method (Project)
keywords: vbapj.chm1512
f1_keywords:
- vbapj.chm1512
ms.prod: project-server
api_name:
- Project.Application.ReassignSelectedAssns
ms.assetid: ab3df7f1-bc36-2b8a-23d7-30ee0387a785
ms.date: 06/08/2017
---


# Application.ReassignSelectedAssns Method (Project)

Reassigns the selected assignments in the Team Planner view.


## Syntax

 _expression_. **ReassignSelectedAssns**( ** _ResourceID_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ResourceUniqueID_|Required|**Long**|Identification number of the resource for the new assignment, or -65535 for unassigned.|

### Return Value

 **Boolean**


## Remarks

The  **ReassignSelectedAssns** method works only with the Team Planner view.

If you use the Team Planner to drag an assignment from one resource to another while you are recording a macro, the macro does not show the results of the drag action. To record a macro that shows the  **ReassignSelectedAssns** method, you must right-click an assignment in the Team Planner, and then click **Reassign To** in the option menu.


## Example

The following line of code reassigns the assignments selected in the Team Planner to the resource with ID = 2.


```
ReassignSelectedAssns ResourceID:=2
```

The following line of code changes the assignments to unassigned.




```
ReassignSelectedAssns ResourceID:=-65535
```


