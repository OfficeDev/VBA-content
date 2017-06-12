---
title: Project.ResourcePoolName Property (Project)
keywords: vbapj.chm132573
f1_keywords:
- vbapj.chm132573
ms.prod: project-server
api_name:
- Project.Project.ResourcePoolName
ms.assetid: 74d426a7-00ed-7a29-5f25-e0f2391add4d
ms.date: 06/08/2017
---


# Project.ResourcePoolName Property (Project)

Gets the name of the enterprise resource pool that a project uses in Project Professional. Read-only  **String**.


## Syntax

 _expression_. **ResourcePoolName**

 _expression_ A variable that represents a **Project** object.


## Remarks

If the project is using enterprise resources,  **ResourcePoolName** gets the name of the virtual resource pool. For example, in Project, the value is "VirtualResPool1".

If the project is not using enterprise resources,  **ResourcePoolName** gets the path and name of the project.


