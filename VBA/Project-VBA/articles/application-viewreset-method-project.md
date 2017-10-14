---
title: Application.ViewReset Method (Project)
keywords: vbapj.chm309
f1_keywords:
- vbapj.chm309
ms.prod: project-server
api_name:
- Project.Application.ViewReset
ms.assetid: ea972480-6417-55a7-9b8e-6cc9944df6c9
ms.date: 06/08/2017
---


# Application.ViewReset Method (Project)

Resets the current view back to the global view definition. 


## Syntax

 _expression_. **ViewReset**

 _expression_ An expression that returns an **Application** object.


### Return Value

 **Boolean**


## Remarks

The  **ViewReset** method displays a dialog box that asks if you are sure you want to continue. Custom formatting, filters, and grouping updates are removed if they are not in the global copy of the view, but the project data is not affected. The **ViewReset** action cannot be undone.

The  **ViewReset** method has the same effect as the **Reset to Default** command in the drop-down list of views on the ribbon.


