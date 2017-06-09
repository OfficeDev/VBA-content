---
title: Application.TimelineShowHide Method (Project)
keywords: vbapj.chm62
f1_keywords:
- vbapj.chm62
ms.prod: project-server
api_name:
- Project.Application.TimelineShowHide
ms.assetid: 237052c0-445b-db78-9a74-10e8742a493d
ms.date: 06/08/2017
---


# Application.TimelineShowHide Method (Project)

Shows or hides the specified feature in the Timeline view.


## Syntax

 _expression_. **TimelineShowHide**( ** _Item_**, ** _Show_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Item_|Required|**PjTimelineShowHide**|Specifies the feature to show or hide. Can be one of the  **[PjTimelineShowHide](pjtimelineshowhide-enumeration-project.md)** constants.|
| _Show_|Optional|**Boolean**|**False** if the feature is hidden; otherwise, **True**. The default value is **True**, which shows the feature.|

### Return Value

 **Boolean**


## Remarks

The  **TimelineShowHide** method corresponds to several commands in the **Show/Hide** group on the **Format** tab on the ribbon. The **Format** tab displays the **Show/Hide** group when the Timeline view is selected. If the Timeline view is not selected, the **TimelineShowHide** method results in error 1100, "The method is not available in this situation."


## Example

The following statement hides the time scale on the timeline.


```vb
Application.TimelineShowHide Item:=pjTimelineShowHideTimescale, Show:=False
```


