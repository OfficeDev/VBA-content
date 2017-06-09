---
title: Resource.ActualOvertimeWork Property (Project)
ms.prod: project-server
api_name:
- Project.Resource.ActualOvertimeWork
ms.assetid: 1770bb0b-8a32-0af6-ddd9-5047b09e4e26
ms.date: 06/08/2017
---


# Resource.ActualOvertimeWork Property (Project)

Gets the actual overtime work (in minutes) for a resource. Read-only  **Variant**.


## Syntax

 _expression_. **ActualOvertimeWork**

 _expression_ A variable that represents a **Resource** object.


## Remarks

The  **ActualOvertimeWork** property does not return any meaningful information for material resources. Setting a value returns a trappable error (error code 1101) when applied to material resources.


