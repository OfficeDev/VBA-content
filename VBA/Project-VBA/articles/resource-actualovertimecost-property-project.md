---
title: Resource.ActualOvertimeCost Property (Project)
ms.prod: project-server
api_name:
- Project.Resource.ActualOvertimeCost
ms.assetid: 9a8579b6-a3ee-7041-98ad-b28adfc51bfc
ms.date: 06/08/2017
---


# Resource.ActualOvertimeCost Property (Project)

Gets the actual overtime cost for a resource. Read-only  **Variant**.


## Syntax

 _expression_. **ActualOvertimeCost**

 _expression_ A variable that represents a **Resource** object.


## Remarks

The  **ActualOvertimeCost** property does not return any meaningful information for material resources. Setting a value returns a trappable error (error code 1101) when applied to material resources.


