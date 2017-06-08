---
title: Application.TaskOnTimelineEx Method (Project)
keywords: vbapj.chm159
f1_keywords:
- vbapj.chm159
ms.assetid: 4307f842-0ccc-d7ac-f386-ec8d259011c6
ms.date: 06/08/2017
ms.prod: project-server
---


# Application.TaskOnTimelineEx Method (Project)

Manages tasks on the Timeline pane or for a specified custom timeline, including specifying the bar that you want to add or remove. Introduced in Office 2016.


## Syntax

 _expression_. **TaskOnTimelineEx**( _TaskID_,  _TaskID_,  _Remove_,  _TimelineViewName_,  _ShowDialog_,  _BarIndex_)

 _expression_ A variable that represents a **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _TaskID_|Optional|**Long**|Specifies the identification number of a task to add to the timeline or remove from the timeline. If  _TaskID_ is specified, selected tasks are ignored.|
| _Remove_|Optional|**Boolean**|**True** if the task specified by _TaskID_ or the selected tasks are removed from the timeline; otherwise, **False**. The default value is **False**.|
| _TimelineViewName_|Optional|**String**|Specifies the name of a timeline to use. The name can be the built-in "Timeline" or an existing custom timeline such as "My Timeline". The default value is the name of the active timeline.|
| _ShowDialog_|Optional|**Boolean**|**True** if Project displays the **Add Tasks to Timeline** dialog box; otherwise, **False**. Any tasks that are already on the timeline have a check by their names. If _ShowDialog_ is **True**, Project ignores the _TaskID_ and _Remove_ arguments. The default value is **False**.|
| _BarIndex_|Optional|**Variant**|The bar that you want to add or remove.|

### Return Value

 **BOOL**


