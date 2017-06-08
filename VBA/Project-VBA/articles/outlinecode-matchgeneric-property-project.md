---
title: OutlineCode.MatchGeneric Property (Project)
ms.prod: project-server
api_name:
- Project.OutlineCode.MatchGeneric
ms.assetid: 5c724bc3-0d2c-8fdc-1f5e-4b62a7d3f761
ms.date: 06/08/2017
---


# OutlineCode.MatchGeneric Property (Project)

 **True** if Project uses the enterprise text custom field (which is equivalent to an outline code) in the Resource Substitution Wizard. Read/write **Boolean**.


## Syntax

 _expression_. **MatchGeneric**

 _expression_ A variable that represents an **OutlineCode** object.


## Remarks

If there are no values in the enterprise lookup table, then the  **MatchGeneric** property is **False** and non-writeable. For local outline codes, **MatchGeneric** is always **False** and non-writeable.


