---
title: Application.GanttBarLinks Method (Project)
keywords: vbapj.chm2071
f1_keywords:
- vbapj.chm2071
ms.prod: project-server
api_name:
- Project.Application.GanttBarLinks
ms.assetid: 80f8fdaa-e08f-3c5e-64dc-43d3dccd7f86
ms.date: 06/08/2017
---


# Application.GanttBarLinks Method (Project)

Shows or hides task links on the Gantt Chart.


## Syntax

 _expression_. **GanttBarLinks**( ** _Display_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Display_|Optional|**Long**|Where links will be drawn from the ends of predecessor links. Can be one of the  **[PjGanttBarLink](pjganttbarlink-enumeration-project.md)** constants. The default value is **PjNoGanttBarLinks**.|

### Return Value

 **Boolean**


## Example

The following example first clears the links and then displays them from the end of one task bar to the top of the next task bar.


```vb
Sub GanttBar_Links() 
'First clear links, than links from end to top of the next bar 
 'Activate Gantt Chart view 
 ViewApply Name:="&;Gantt Chart" 
 GanttBarLinks pjNoGanttBarLinks 
 GanttBarLinks pjToTop 
End Sub
```


