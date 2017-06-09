---
title: Application.ViewCopy Method (Project)
keywords: vbapj.chm300
f1_keywords:
- vbapj.chm300
ms.prod: project-server
api_name:
- Project.Application.ViewCopy
ms.assetid: b1ed6b3e-ad95-15f4-80bd-054d608ef9a1
ms.date: 06/08/2017
---


# Application.ViewCopy Method (Project)

Copies the current view.


## Syntax

 _expression_. **ViewCopy**( ** _Name_**, ** _ApplyTo_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**String**|Name of the copy of the view.|
| _ApplyTo_|Optional|**Integer**|Specifies which pane of a split view is copied. The value can be one of the following:
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




 **Note**  In a combination view, the primary pane is the view that remains when a details or secondary pane is closed. Usually the primary pane is at the top; however, the Timeline is a secondary pane, but it displays at the top. For example, with the Resource Sheet view, clicking  **Details** on the **View** tab on the Ribbon shows the secondary Resource Form pane on the bottom. Clicking **Timeline** closes the Resource Form at the bottom and opens the Timeline pane at the top.

Using the  **ViewCopy** method with no arguments displays the **Save View** dialog box, which enables the user to name the copy of the view.


## Example

If the current view includes the Timeline in the top pane and the Gantt Chart in the bottom pane, where the Gantt Chart is the active pane, the following statement copies the Timeline view. After you execute the statement, the drop-down list of views includes  **Copy of Timeline** in the **Custom** section.


```
application.ViewCopy Name:="Copy of Timeline", ApplyTo:=1
```


