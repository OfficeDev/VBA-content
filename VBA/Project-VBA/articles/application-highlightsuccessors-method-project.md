---
title: Application.HighlightSuccessors Method (Project)
keywords: vbapj.chm149
f1_keywords:
- vbapj.chm149
ms.prod: project-server
ms.assetid: 7a72cc0a-49f0-c95d-23cc-35d7ee077539
ms.date: 06/08/2017
---


# Application.HighlightSuccessors Method (Project)
Sets or clears task successor highlighting for the task path feature.

## Syntax

 _expression_. **HighlightSuccessors** _(Set)_

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Set_|Optional|**Variant**|**True** to set task successor highlighting; **False** to clear task successor highlighting.|
| _Set_|Optional|VARIANT||
|Name|Required/Optional|Data type|Description|

### Return value

 **Boolean**


## Remarks

The  **HighlightSuccessors** method corresponds to the **Successors** item in the **Task Path** drop-down list, on the **FORMAT** tab, under **GANTT CHART TOOLS** on the ribbon.


## Example

Create a project where task 4 is a successor of task 3, and then run the following statements in the  **Immediate** window of the VBE. The **PathSuccessor** statement prints **True**.


```vb
Application.SelectRow Row:=3, RowRelative:=False 
Application.HighlightSuccessors True
? ActiveProject.Tasks(4).PathSuccessor

```


## See also


#### Concepts


[Application Object](application-object-project.md)
#### Other resources


[Task.PathSuccessor Property](task-pathsuccessor-property-project.md)
