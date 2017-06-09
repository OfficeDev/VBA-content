---
title: Application.GanttBarStyleBaseline Method (Project)
keywords: vbapj.chm83
f1_keywords:
- vbapj.chm83
ms.prod: project-server
api_name:
- Project.Application.GanttBarStyleBaseline
ms.assetid: c9cb0ebb-998c-c9ea-9d3f-5cb06813c364
ms.date: 06/08/2017
---


# Application.GanttBarStyleBaseline Method (Project)

Shows or hides the specified baseline on Gantt bars of the active view.


## Syntax

 _expression_. **GanttBarStyleBaseline**( ** _Baseline_**, ** _Show_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Baseline_|Required|**Integer**|Specifies the baseline number. Valid values are 0 through 10.|
| _Show_|Required|**Boolean**|If  **True**, show the baseline. If **False**, hide the baseline.|

### Return Value

 **Boolean**


## Remarks

On the Ribbon, the  **GanttBarStyleBaseline** method corresponds to the **Baseline** drop-down list in the **Bar Styles** group on the **Format** tab for **Gantt Chart Tools**.


## Example

The following command shows the baseline 1 data for tasks on the active Gantt chart view.


```
GanttBarStyleBaseline(1, True)
```


