---
title: Application.SidepaneToggle Method (Project)
keywords: vbapj.chm52
f1_keywords:
- vbapj.chm52
ms.prod: project-server
api_name:
- Project.Application.SidepaneToggle
ms.assetid: 882c9bef-f150-7128-a506-388dbe39558d
ms.date: 06/08/2017
---


# Application.SidepaneToggle Method (Project)

Triggers the  **[WindowSidepaneDisplayChange](application-windowsidepanedisplaychange-event-project.md)** event, which shows or hides the side pane of the **Project Guide**. Deprecated in Project.


## Syntax

 _expression_. **SidepaneToggle**( ** _Show_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Show_|Optional|**Boolean**|**True** if Project shows the side pane for the **Project Guide**.  **False** if Project hides the side pane for the **Project Guide**.|

### Return Value

 **Boolean**


## Remarks

The  **SidepaneToggle** method is used to change the side pane display state; you cannot use this method to return the current display state of the side pane in the **Project Guide**.


 **Note**  The Project Guide is disabled by default in Project. Although you can create and display custom Project Guide pages, we recommend that you create a task pane app instead of the Project Guide for new development.


