---
title: Application.ResourceComparison Method (Project)
keywords: vbapj.chm2185
f1_keywords:
- vbapj.chm2185
ms.prod: project-server
api_name:
- Project.Application.ResourceComparison
ms.assetid: 42223a8d-cc71-26c0-35e8-c184b40a46c2
ms.date: 06/08/2017
---


# Application.ResourceComparison Method (Project)

In a project comparison report, shows the Resource Sheet view in all three project plans, to compare resources.


## Syntax

 _expression_. **ResourceComparison**

 _expression_ An expression that returns an **Application** object.


### Return Value

 **Boolean**


## Remarks

Use the  **CompareProjectVersions** method to create a project comparison report, or choose **Compare Projects** on the **PROJECT** ribbon.

The  **ResourceComparison** method is equivalent to the **Resource Comparison** command on the **Compare Projects** tab on the ribbon.

To compare tasks in a comparison report, use the  **[TaskComparison](application-taskcomparison-method-project.md)** method.


