---
title: Application.GoToItemInVersions Method (Project)
keywords: vbapj.chm2186
f1_keywords:
- vbapj.chm2186
ms.prod: project-server
api_name:
- Project.Application.GoToItemInVersions
ms.assetid: 51b7e580-978d-17cc-f293-bb30d77c48c2
ms.date: 06/08/2017
---


# Application.GoToItemInVersions Method (Project)

For the selected item in a project version comparison report, highlights that item in each version.


## Syntax

 _expression_. **GoToItemInVersions**

 _expression_ An expression that returns an **Application** object.


### Return Value

 **Boolean**


## Remarks

When you compare two versions of a project file, Project creates a new project named  **Comparison Report** and shows each of the original versions below the **Comparison Report** window. If an item is selected in the **Comparison Report** window, **GoToItemInVersions** selects the same item in each of the original versions. Focus changes to the second version window.

The  **GoToItemInVersions** method is equivalent to the **Go to Item** command in the **Compare** group of the **Compare Projects** tab on the Ribbon.


