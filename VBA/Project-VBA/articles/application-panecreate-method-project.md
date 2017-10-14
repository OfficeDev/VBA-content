---
title: Application.PaneCreate Method (Project)
keywords: vbapj.chm2003
f1_keywords:
- vbapj.chm2003
ms.prod: project-server
api_name:
- Project.Application.PaneCreate
ms.assetid: 6ecf7151-eaeb-4a28-c877-a6e5366e2a8e
ms.date: 06/08/2017
---


# Application.PaneCreate Method (Project)

Creates a lower pane for the active window.


## Syntax

 _expression_. **PaneCreate**

 _expression_ A variable that represents an **Application** object.


### Return Value

 **Boolean**


## Remarks

If the active view is one of the task views, including the Task Usage view, the new pane will be the Task Form. If the active view is one of the resource views, including the Resource Usage view, the new pane will be the Resource Form.


