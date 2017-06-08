---
title: Application.TimelineFormat Method (Project)
keywords: vbapj.chm64
f1_keywords:
- vbapj.chm64
ms.prod: project-server
api_name:
- Project.Application.TimelineFormat
ms.assetid: 96f936a1-15be-8df4-4683-cd876c8a69ce
ms.date: 06/08/2017
---


# Application.TimelineFormat Method (Project)

Formats the  **Timeline** view to specify the number of text lines in timeline tasks and whether to show or hide details.


## Syntax

 _expression_. **TimelineFormat**( ** _NumLines_**, ** _Minimized_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _NumLines_|Optional|**Variant**|Number of text lines to show in tasks on the timeline. Values can be 1 through 10; other values are ignored.|
| _Minimized_|Optional|**Boolean**|If  **true**, minimizes the timeline so that tasks do not show details. If **false**, vertically expands the timeline so that tasks show detail text lines.|

### Return Value

 **Boolean**


## Remarks

The  **TimelineFormat** method parameters correspond to the **Text Lines** command and the **Detailed Timeline** command on the **Format** tab for **Timeline Tools** in the ribbon.

If the timeline does not show tasks, the  _Minimized_ parameter has no effect. If the **Timeline** view is not active, the **TimelineFormat** method results in run-time error 1100, "Application-defined or object-defined error."


## Example

If the timeline shows one or more tasks, the following example sets four text lines in each task and expands the timeline to show all four lines.


```vb
Sub FormatTimeline() 
    TimelineFormat NumLines:=4, Minimized:=False 
End Sub
```


