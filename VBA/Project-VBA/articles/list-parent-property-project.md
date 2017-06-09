---
title: List.Parent Property (Project)
ms.prod: project-server
api_name:
- Project.List.Parent
ms.assetid: 08d2d7d8-fafc-8f60-be78-c2d462005eaf
ms.date: 06/08/2017
---


# List.Parent Property (Project)

Gets the parent of the  **List** object. Read-only **Object**.


## Syntax

 _expression_. **Parent**

 _expression_ A variable that represents a **List** object.


## Remarks

The parent of a  **List** object can be a **Selection** (with the **FieldIDList** and **FieldNameList** properties), a **Project** (including several properties such as **MapList**, **ReportList**, and **ViewList** ).

Use the  **Parent** property to access the properties or methods of the parent of an object.


