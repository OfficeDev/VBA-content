---
title: Application.ViewEditSingle Method (Project)
keywords: vbapj.chm303
f1_keywords:
- vbapj.chm303
ms.prod: project-server
api_name:
- Project.Application.ViewEditSingle
ms.assetid: 445977e9-e540-14b3-a179-ea132491265e
ms.date: 06/08/2017
---


# Application.ViewEditSingle Method (Project)

Creates, edits, or copies a single-pane view.


## Syntax

 _expression_. **ViewEditSingle**( ** _Name_**, ** _Create_**, ** _NewName_**, ** _Screen_**, ** _ShowInMenu_**, ** _HighlightFilter_**, ** _Table_**, ** _Filter_**, ** _Group_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**String**| The name of a single-pane view to edit, create, or copy. The default is the name of the active view.|
| _Create_|Optional|**Boolean**|**True** if Project creates a single-pane view. If NewName is an empty string (""), the new view is given the name specified with Name. Otherwise, the new view is a copy of the view specified with Name and is given the name specified with NewName. The default value is **False**.|
| _NewName_|Optional|**String**|A new name for the view specified with Name (Create is  **False** ) or a name for the new view just created (Create is **True** ). If NewName is an empty string and Create is **False**, the view specified with Name retains its current name. The default value is **False**.|
| _Screen_|Optional|**Long**|A constant specifying the view to display. Can be one of the  **[PjViewScreen](pjviewscreen-enumeration-project.md)** constants. The default value is **pjGantt**|
| _ShowInMenu_|Optional|**Boolean**|**True** if the view name appears on the **Other Views** drop-down menu. The default value is **False**.|
| _HighlightFilter_|Optional|**Boolean**|**True** if Project should highlight filtered items. The default value is **False**.|
| _Table_|Optional|**String**|The name of a table to display in the view. Required for a new view.|
| _Filter_|Optional|**String**|The name of a filter to apply to the view. Required for a new view.|
| _Group_|Optional|**String**|The name of a group to apply to the view. If a group is required for the view, but none is specified, the default value is "No Group". The Group value is ignored if the view specified with the Screen argument does not use groups.|

### Return Value

 **Boolean**


## Example

The following example creates a new view for tasks currently in progress and grouped by duration.


```vb
Sub DisplayMyTasks() 
 ViewEditSingle Name:="My Tasks", Create:=True, _ 
 Screen:=pjGantt, Table:="Schedule", _ 
 Filter:="In Progress Tasks", Group:="Duration" 
End Sub
```


