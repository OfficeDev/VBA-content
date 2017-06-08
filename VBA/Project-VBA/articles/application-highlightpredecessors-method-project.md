---
title: Application.HighlightPredecessors Method (Project)
keywords: vbapj.chm147
f1_keywords:
- vbapj.chm147
ms.prod: project-server
ms.assetid: e4c51516-2e5d-3ef9-3165-84fe6f9ad38b
ms.date: 06/08/2017
---


# Application.HighlightPredecessors Method (Project)
Sets or clears task predecessor highlighting for the task path feature.

## Syntax

 _expression_. **HighlightPredecessors** _(Set)_

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Set_|Optional|**Variant**|**True** to set task predecessor highlighting; **False** to clear the task predecessor highlighting.|
| _Set_|Optional|VARIANT||
|Name|Required/Optional|Data type|Description|

### Return value

 **Boolean**


## Remarks

The  **HighlightPredecessors** method corresponds to the ** Predecessors** item in the **Task Path** drop-down list, on the **FORMAT** tab, under **GANTT CHART TOOLS** on the ribbon.


## Example

Create a project where task 2 is a predecessor of task 3, and then run the following statements in the  **Immediate** window of the VBE. The **PathPredecessor** statement prints **True**.


```vb
Application.SelectRow Row:=2, RowRelative:=False 
Application.HighlightPredecessors True
? ActiveProject.Tasks(3).PathPredecessor
```


## See also


#### Concepts


[Application Object](application-object-project.md)
#### Other resources


[Task.PathPredecessor Property](task-pathpredecessor-property-project.md)
