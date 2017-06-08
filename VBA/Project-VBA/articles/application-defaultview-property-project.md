---
title: Application.DefaultView Property (Project)
ms.prod: project-server
api_name:
- Project.Application.DefaultView
ms.assetid: 19f3cc23-6267-0b1f-7db5-7783d6936533
ms.date: 06/08/2017
---


# Application.DefaultView Property (Project)

Gets or sets the name of the view that appears when you start Project. Read/write  **String**.


## Syntax

 _expression_. **DefaultView**

 _expression_ A variable that represents an **Application** object.


## Remarks

The  **DefaultView** property can be the name of a custom view or one of the following built-in views:


|||
|:-----|:-----|
|"Bar Rollup"|"Resource Graph"|
|"Calendar"|"Resource Name Form"|
|"Descriptive Network Diagram"|"Resource Sheet"|
|"Detail Gantt"|"Resource Usage"|
|"Gantt Chart"|"Task Details Form"|
|"Leveling Gantt"|"Task Entry"|
|"Milestone Date Rollup"|"Task Form"|
|"Milestone Rollup"|"Task Name Form"|
|"Multiple Baselines Gantt"|"Task Sheet"|
|"Network Diagram"|"Task Usage"|
|"Relationship Diagram"|"Team Planner"|
|"Resource Allocation"|"Timeline"|
|"Resource Form"|"Tracking Gantt"|
The default value is "Gantt with Timeline".


