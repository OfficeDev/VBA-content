---
title: Application.PanZoomPanTo Method (Project)
ms.prod: project-server
api_name:
- Project.Application.PanZoomPanTo
ms.assetid: 7bdca9f2-d006-6cab-872b-01cf54f6e8ce
ms.date: 06/08/2017
---


# Application.PanZoomPanTo Method (Project)

Pans the Gantt chart in the active view to the specified start date.


## Syntax

 _expression_. **PanZoomPanTo**( ** _Start_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Start_|Required|**Variant**|Specifies the start date for the left side of the Gantt chart.|

### Return Value

Nothing


## Remarks

The  **PanZoomPanTo** method has no effect on the Calendar view or Network Diagram (PERT chart) view.

To zoom the Gantt chart in or out, which changes the timescale, use the  **[PanZoomZoomTo](application-panzoomzoomto-method-project.md)** method. To change the timescale format and labels, use the **[TimescaleEdit](application-timescaleedit-method-project.md)** method.


## Example

The following command moves the beginning of the visible Gantt chart to March 18, 2012.


```
PanZoomPanTo Start:="3/18/2012" 

```


