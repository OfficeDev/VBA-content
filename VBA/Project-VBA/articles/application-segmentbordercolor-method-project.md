---
title: Application.SegmentBorderColor Method (Project)
keywords: vbapj.chm72
f1_keywords:
- vbapj.chm72
ms.prod: project-server
api_name:
- Project.Application.SegmentBorderColor
ms.assetid: 99c2d2ba-f0c5-b462-5801-ac9c7ee75a02
ms.date: 06/08/2017
---


# Application.SegmentBorderColor Method (Project)

Sets the border color for the assignment segments of a selected task in the Team Planner view.


## Syntax

 _expression_. **SegmentBorderColor**( ** _Color_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Color_|Required|**Long**|Border color of the assignment segments. The color is a hexadecimal RGB value, where red is the last byte.|

### Return Value

 **Boolean**


## Example

In the following example, a task is assigned to two resources. After selecting either of the assignments, running the  **ChangeSegmentColor** macro shows all assignments for the task as light red with a blue border.


```vb
Sub ChangeSegmentColor() 
    Application.SegmentFillColor(&;H8080FF) 
    Application.SegmentBorderColor(&;HFF1010) 
End Sub
```


