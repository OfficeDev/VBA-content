---
title: Application.SetSplitBar Method (Project)
keywords: vbapj.chm31
f1_keywords:
- vbapj.chm31
ms.prod: project-server
api_name:
- Project.Application.SetSplitBar
ms.assetid: caf26a56-43ad-1714-79e4-cab013a55f3c
ms.date: 06/08/2017
---


# Application.SetSplitBar Method (Project)

Positions the vertical split bar in a sheet view to display the specified number of columns.


## Syntax

 _expression_. **SetSplitBar**( ** _ShowColumns_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ShowColumns_|Optional|**Long**|Specifies the number of columns to display, including the locked  **ID** column. The value can be 1 through 75. The default value is the total number of columns currently displayed, including any partially hidden column.|

### Return Value

 **Boolean**


## Remarks

If the right pane of the view has been sized so that there is no left pane, the  **SetSplitBar** method has no effect.

If the split bar is partially hiding the right-most column, using the  **SetSplitBar** method with no argument moves the split bar to show the complete column.


## Example

If the standard Gantt Chart is the active view, the following statement sets the split bar to show only the  **ID**,  **Indicators**,  **Task Mode**, and  **Task Name** columns.


```vb
Application.SetSplitBar ShowColumns:=4
```


