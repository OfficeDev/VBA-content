---
title: Resource.EnterpriseUniqueID Property (Project)
keywords: vbapj.chm132200
f1_keywords:
- vbapj.chm132200
ms.prod: project-server
api_name:
- Project.Resource.EnterpriseUniqueID
ms.assetid: ad5bdf09-a1e0-c9fd-c3ae-ba1639177a95
ms.date: 06/08/2017
---


# Resource.EnterpriseUniqueID Property (Project)

Gets the enterprise unique identification number for a resource. Read-only  **Long**.


## Syntax

 _expression_. **EnterpriseUniqueID**

 _expression_ A variable that represents a **Resource** object.


## Remarks

The enterprise unique ID is a unique number for the enterprise resource within an instance of Project Web App. For example, the first enterprise resource defined has the unique ID 1, the second enterprise resource is 2, and so forth. The  **Guid** property is the only absolutely unique identification for a resource. For local resources in an enterprise project, the **EnterpriseUniqueID** value is -1.

The  **EnterpriseUniqueID** property is available only in Project Professional.


