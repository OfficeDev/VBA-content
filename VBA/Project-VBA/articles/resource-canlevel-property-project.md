---
title: Resource.CanLevel Property (Project)
ms.prod: project-server
api_name:
- Project.Resource.CanLevel
ms.assetid: 21d1f14d-4d53-21b4-a164-cf6ab9e2cf98
ms.date: 06/08/2017
---


# Resource.CanLevel Property (Project)

 **True** if the resource can be leveled. Read/write **Variant**.


## Syntax

 _expression_. **CanLevel**

 _expression_ An expression that returns a **Resource** object.


## Remarks

The  **CanLevel** property does not return any meaningful information for material resources. Setting a value returns a trappable error (error code 1101) when applied to material resources.


