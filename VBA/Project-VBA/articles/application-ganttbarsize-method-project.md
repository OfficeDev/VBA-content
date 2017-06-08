---
title: Application.GanttBarSize Method (Project)
keywords: vbapj.chm2058
f1_keywords:
- vbapj.chm2058
ms.prod: project-server
api_name:
- Project.Application.GanttBarSize
ms.assetid: 691ee987-a62b-bf5f-0088-0f153aa64966
ms.date: 06/08/2017
---


# Application.GanttBarSize Method (Project)

Sets the height, in points, of the Gantt bars in the active Gantt Chart.


## Syntax

 _expression_. **GanttBarSize**( ** _Size_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Size_|Required|**Long**|A constant specifying the height, in points, of the Gantt bars in the active Gantt Chart. Can be one of the following  **[PjBarSize](pjbarsize-enumeration-project.md)** constants.|

### Return Value

 **Boolean**


## Example

The following example set the bar size to pjBarSize24.


```vb
Sub GanttBar_Size() 
 
 'Activate Gantt Chart view 
 ViewApply Name:="&;Gantt Chart" 
 GanttBarSize Size:= 
pjBarSize24
```


```
End Sub
```


