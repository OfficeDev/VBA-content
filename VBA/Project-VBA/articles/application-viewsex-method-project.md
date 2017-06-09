---
title: Application.ViewsEx Method (Project)
keywords: vbapj.chm310
f1_keywords:
- vbapj.chm310
ms.prod: project-server
api_name:
- Project.Application.ViewsEx
ms.assetid: 42567343-54df-fbf2-64a3-79ba72d12866
ms.date: 06/08/2017
---


# Application.ViewsEx Method (Project)

Displays the  **More Views** dialog box with the specified pane of the current view selected, which prompts the user to manage views.


## Syntax

 _expression_. **ViewsEx**( ** _ApplyTo_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ApplyTo_|Optional|**Integer**|Specifies which pane of a split view is selected. The value can be one of the following:
|||
|:-----|:-----|
|**Value**|**Description**|
|0|Primary (usually the top) pane of a split view|
|1|Secondary (usually the bottom) pane of a split view|
|4|Active pane|
|5|Primary pane, or the Timeline if it is active|
|

### Return Value

 **Boolean**


## Remarks




 **Note**  In a combination view, the primary pane is the view that remains when a details or secondary pane is closed. Usually the primary pane is at the top; however, the Timeline is a secondary pane, but it displays at the top. For example, with the Resource Sheet view, clicking  **Details** on the **View** tab of the Ribbon shows the secondary Resource Form pane on the bottom. Clicking **Timeline** closes the Resource Form at the bottom and opens the Timeline pane at the top.


## Example

If the current view includes the Timeline and the Gantt Chart, where the Timeline is the active pane, the following example shows  **Timeline** selected in the **More Views** dialog box.


```
application.ViewsEx ApplyTo:=5
```


