---
title: Application.TaskOnTimeline Method (Project)
keywords: vbapj.chm60
f1_keywords:
- vbapj.chm60
ms.prod: project-server
api_name:
- Project.Application.TaskOnTimeline
ms.assetid: 8201380b-f0ae-4e53-7461-e323ad6fe5e2
ms.date: 06/08/2017
---


# Application.TaskOnTimeline Method (Project)

Manages tasks on the Timeline pane or for a specified custom timeline.


## Syntax

 _expression_. **TaskOnTimeline**( ** _TaskID_**, ** _Remove_**, ** _TimelineViewName_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _TaskID_|Optional|**Long**|Specifies the identification number of a task to add to the timeline or remove from the timeline. If  _TaskID_ is specified, selected tasks are ignored.|
| _Remove_|Optional|**Boolean**|**True** if the task specified by _TaskID_ or the selected tasks are removed from the timeline; otherwise, **False**. The default value is **False**.|
| _TimelineViewName_|Optional|**String**|Specifies the name of a timeline to use. The name can be the built-in "Timeline" or an existing custom timeline such as "My Timeline". The default value is the name of the active timeline.|
| _ShowDialog_|Optional|**Boolean**|**True** if Project displays the **Add Tasks to Timeline** dialog box; otherwise, **False**. Any tasks that are already on the timeline have a check by their names. If _ShowDialog_ is **True**, Project ignores the _TaskID_ and _Remove_ arguments. The default value is **False**.|

### Return Value

 **Boolean**


## Remarks

Running  **TaskOnTimeline** with no arguments puts selected tasks on the active timeline.

If the project includes custom timeline views, using the  _TimelineViewName_ argument activates the specified timeline, and then applies changes specified by the other arguments. If the specified timeline does not exist, **TaskOnTimeline** takes no action, but still returns **True**.


## Example

The following statement removes selected tasks from the timeline. You can select the tasks in the Gantt Chart or on the timeline.


```
application.TaskOnTimeline Remove:=True
```

If the built-in Timeline pane is active and a custom timeline named "My Timeline" exists, the following statement replaces the Timeline pane with "My Timeline", and then adds task 3 to the custom timeline.




```
application.TaskOnTimeline TaskID:=3, TimelineViewName:="My Timeline"
```


