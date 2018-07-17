---
title: Application.CopyReport Method (Project)
keywords: vbapj.chm141
f1_keywords:
- vbapj.chm141
ms.prod: project-server
ms.assetid: 9f1e59d5-a2a5-4c8f-1c01-b1c63046558d
ms.date: 06/08/2017
---


# Application.CopyReport Method (Project)
Makes a copy of the active report to the clipboard.

## Syntax

 _expression_. **CopyReport**

 _expression_ A variable that represents an **Application** object.


### Return value

 **Boolean**

 **True** if the **CopyReport** method is successful.


## Remarks

You can paste the copied report into another application, such as Word, Excel, or PowerPoint. The  **CopyReport** method corresponds to the **Copy Report** command on the **DESIGN** tab of the **REPORT TOOLS** ribbon.

The  **CopyReport** method does not apply to views, such as the following:


- Calendar
    
- Gantt Chart
    
- PERT Chart (Network Diagram)
    
- Resource Form
    
- Resource Sheet
    
- Resource histogram
    
- Resource Usage
    
- Task Form
    
- Task Sheet
    
- Task Usage
    
- Timeline
    
If you use the  **CopyReport** method on a view that is not supported, Project shows run-time error 1100, **Application-defined or object-defined error**.


## See also


#### Other resources


[Reports.Item](reports-item-property-project.md)
