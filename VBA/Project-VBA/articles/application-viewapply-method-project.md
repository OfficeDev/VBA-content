---
title: Application.ViewApply Method (Project)
keywords: vbapj.chm302
f1_keywords:
- vbapj.chm302
ms.prod: project-server
api_name:
- Project.Application.ViewApply
ms.assetid: 3e0d3fbd-5aa7-ceb8-b926-79646986d464
ms.date: 06/08/2017
---


# Application.ViewApply Method (Project)

Applies a view to the active window.


## Syntax

 _expression_. **ViewApply**( ** _Name_**, ** _SinglePane_**, ** _Toggle_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**String**|The name of the view to display in the active window.|
| _SinglePane_|Optional|**Boolean**|**True** if an existing split is removed and the active window displays a single-pane view. The default value is **False**.|
| _Toggle_|Optional|**Boolean**|**True** if the active window switches from one pane to two panes, or from two panes to one pane. Toggle is ignored if SinglePane is **True**. The default value is **False**.|

### Return Value

 **Boolean**


## Remarks

To apply a view where the change can be one of the built-in views and can be applied to a specified pane in a split view, use the  **[ViewApplyEx](application-viewapplyex-method-project.md)** method.


## Example

The following example sets the active window to a single-pane view of the Resource Sheet. It assumes that the active view is a combination of the Gantt Chart and the Task Form details view.


```vb
Sub ChangeWindowToResourceSheet() 
 ViewApply Name:="Resource Sheet", SinglePane:=True 
End Sub
```


