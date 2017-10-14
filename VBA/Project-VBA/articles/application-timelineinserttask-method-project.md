---
title: Application.TimelineInsertTask Method (Project)
keywords: vbapj.chm65
f1_keywords:
- vbapj.chm65
ms.prod: project-server
api_name:
- Project.Application.TimelineInsertTask
ms.assetid: 4a1833a4-ddbb-577d-fe58-5907644fd127
ms.date: 06/08/2017
---


# Application.TimelineInsertTask Method (Project)

When the Timeline view is selected, displays the  **Task Information** dialog box, and then inserts a new task into the project and adds the task to the Timeline view.


## Syntax

 _expression_. **TimelineInsertTask**( ** _Type_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Required|**PjTimelineInsertTaskType**|Specifies the type of task; that is, whether the task is a regular task, a milestone, or a callout task. Can be one of the following  **[PjTimelineInsertTaskType](pjtimelineinserttasktype-enumeration-project.md)** constants: **pjTimelineInsertTask**, **pjTimelineInsertMilestone**, or **pjTimelineInsertCalloutTask**. Any of the task types can be manually or automatically scheduled.|

### Return Value

 **Boolean**


## Remarks

The  **TimelineInsertTask** method shows a manually scheduled or automatically scheduled task in the **Task Information** dialog box, depending on the type of task shown in the **New Tasks** section of the Project status bar.

If the user cancels the  **Task Information** dialog box, **TimelineInsertTask** returns **False**.


 **Note**  The  ** Display on Timeline** check box in the **Task Information** dialog box is clear. The **TimelineInsertTask** method adds a task to the timeline whether the check box is checked or clear.

The  **TimelineInsertTask** method corresponds to the **Task**,  **Callout Task**, and  **Milestone** commands in the **Insert** group on the **Format** tab on the ribbon. The **Format** tab displays the **Insert** group when the Timeline view is selected. If the Timeline view is not selected, the **TimelineInsertTask** method results in error 1100, "The method is not available in this situation."


## Example

If the Project status bar shows  **New Tasks: Manually Scheduled**, the following statement displays the  **Task Information** dialog box, which prompts the user to name a manually scheduled task. The default start date is the project start date. When the user clicks **OK**, Project inserts the task in the Gantt chart and shows the new task on the timeline, with the task information in a callout box attached to the timeline.


```vb
Application.TimelineInsertTask Type:=pjTimelineInsertCalloutTask
```


