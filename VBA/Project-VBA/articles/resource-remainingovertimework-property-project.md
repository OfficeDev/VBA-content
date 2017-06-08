---
title: Resource.RemainingOvertimeWork Property (Project)
ms.prod: project-server
api_name:
- Project.Resource.RemainingOvertimeWork
ms.assetid: f5b3ae63-5983-60e4-517b-b484b35505c0
ms.date: 06/08/2017
---


# Resource.RemainingOvertimeWork Property (Project)

Gets the remaining overtime work (in minutes) for the resource. Read-only  **Variant**.


## Syntax

 _expression_. **RemainingOvertimeWork**

 _expression_ A variable that represents a **Resource** object.


## Remarks

The  **RemainingOvertimeWork** property does not return any meaningful information for material resources. Setting a value returns a trappable error (error code 1101) when applied to material resources.


