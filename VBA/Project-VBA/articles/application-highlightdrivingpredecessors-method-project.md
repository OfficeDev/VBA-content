---
title: Application.HighlightDrivingPredecessors Method (Project)
keywords: vbapj.chm148
f1_keywords:
- vbapj.chm148
ms.prod: project-server
ms.assetid: 2a2653c5-6b7d-9429-f73f-e65c0cda1c5c
ms.date: 06/08/2017
---


# Application.HighlightDrivingPredecessors Method (Project)
Sets or clears task driving predecessor highlighting for the task path feature.

## Syntax

 _expression_. **HighlightDrivingPredecessors** _(Set)_

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Set_|Optional|**Variant**|**True** to set task driving predecessor highlighting; **False** to clear the task driving predecessor highlighting.|
| _Set_|Optional|VARIANT||
|Name|Required/Optional|Data type|Description|

### Return value

 **Boolean**


## Remarks

The  **HighlightDrivingPredecessors** method corresponds to the **Driving Predecessors** item in the **Task Path** drop-down list, on the **FORMAT** tab, under **GANTT CHART TOOLS** on the ribbon.


## Example

Create a project where task 2 is a driving predecessor of task 3, and then run the following statements in the  **Immediate** window of the VBE. The **PathDrivingPredecessor** statement prints **True**.


```vb
Application.SelectRow Row:=2, RowRelative:=False 
Application.HighlightDrivingPredecessors True
? ActiveProject.Tasks(3).PathDrivingPredecessor
```


## See also


#### Concepts


[Application Object](application-object-project.md)
#### Other resources


[Task.PathDrivingPredecessor Property](task-pathdrivingpredecessor-property-project.md)
