---
title: Application.ViewApplyEx Method (Project)
keywords: vbapj.chm311
f1_keywords:
- vbapj.chm311
ms.prod: project-server
api_name:
- Project.Application.ViewApplyEx
ms.assetid: 437ec3b5-d42d-ed79-e8c7-220f797023b5
ms.date: 06/08/2017
---


# Application.ViewApplyEx Method (Project)

Applies a view to the active window, where the change can be one of the built-in views and can be applied to a specified pane in a split view.

## Syntax

_expression_. **ViewApplyEx** (**_Name_**, **_SinglePane_**, **_Toggle_**, **_ApplyTo_**, **_BuiltInView_**)

_expression_ An expression that returns an **Application** object.


### Parameters

|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**String**|The name of the view to display in the active window.|
| _SinglePane_|Optional|**Boolean**|**True** if an existing split is removed and the active window displays a single-pane view. The default value is **False**.|
| _Toggle_|Optional|**Boolean**|**True** if the active window switches from one pane to two panes, or from two panes to one pane. _Toggle_ is ignored if _SinglePane_ is **True**. The default value is **False**.|
| _ApplyTo_|Optional|**Integer**|Specifies where the view is applied. The value can be one of the [ApplyTo values](#applyto-values).|
| _BuiltInView_|Optional|**PjViewType**|Specifies a built-in view. Can be one of the **[PjViewType](pjviewtype-enumeration-project.md)** constants. The default is **pjViewUndefined**. _BuiltInView_ is ignored if _Name_ is specified.|

<br/>

#### ApplyTo values

|||
|:-----|:-----|
|**Value**|**Description**|
|0|Primary (usually the top) pane of a split view|
|1|Secondary (usually the bottom) pane of a split view|
|4|Active pane|
|5|Primary pane, or the Timeline if it is active|

<br/>

### Return value

 **Boolean**

## Remarks

> [!NOTE]
> In a combination view, the primary pane is the view that remains when a details or secondary pane is closed. Usually the primary pane is at the top; however, the Timeline is a secondary pane, but it displays at the top. For example, with the Resource Sheet view, clicking **Details** on the **View** tab on the ribbon shows the secondary Resource Form pane on the bottom. Clicking **Timeline** closes the Resource Form at the bottom and opens the Timeline pane at the top.

The Gantt Chart view cannot be combined with the Team Planner view. Some views, such as the Calendar view, cannot be displayed in a details pane. The **ViewApply** method shows an error message, and then shows error 1004, "An unexpected error occurred with the method."

## Example

The following example sets the single-pane active window to a combination view with the Gantt Chart in the lower pane. It assumes that the active view is something other than the Gantt Chart.

```vb
Sub ChangeWindowToGanttChart() 
    ViewApplyEx Toggle:=True, BuiltInView:=pjViewGantt 
End Sub
```

If the current split view includes the Resource Usage and the Timeline views, where either pane is selected, the following example displays the Resource Usage view in the top pane and the Task Usage view in the bottom pane.

```vb
Sub ChangeSecondaryToTaskForm() 
    ViewApplyEx(Name:="Task Usage", ApplyTo:=1) 
End sub
```

