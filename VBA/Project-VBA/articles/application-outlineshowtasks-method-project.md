---
title: Application.OutlineShowTasks Method (Project)
keywords: vbapj.chm27
f1_keywords:
- vbapj.chm27
ms.prod: project-server
api_name:
- Project.Application.OutlineShowTasks
ms.assetid: 614eb1fc-93eb-3df2-ae52-4fad98c80b3b
ms.date: 06/08/2017
---


# Application.OutlineShowTasks Method (Project)

Expands an outline to show all tasks up to the specified level and collapses any levels below.


## Syntax

 _expression_. **OutlineShowTasks**( ** _OutlineNumber_**, ** _ExpandInsertedProjects_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _OutlineNumber_|Optional|**Long**|The outline level to display. Higher outline levels are expanded to show this level. The level specified with  **OutlineNumber** and lower (if any) are collapsed. Can be one of the **[PjTaskOutlineShowLevel](pjtaskoutlineshowlevel-enumeration-project.md)** constants.|
| _ExpandInsertedProjects_|Optional|**Boolean**|**True** if tasks from subprojects are affected by the value specified with **OutlineNumber**. The default value is **False**.|

### Return Value

 **Boolean**


## Example

This example has the same effect as collapsing the entire outline, including any tasks from subprojects.


```vb
Sub CollapseOutline() 
 Application.OutlineShowTasks pjTaskOutlineShowLevel1, True 
End Sub
```


