---
title: Calendar.BaseCalendar Property (Project)
ms.prod: project-server
api_name:
- Project.Calendar.BaseCalendar
ms.assetid: 3ea2b0e2-8d73-b564-fdd1-a098a8428562
ms.date: 06/08/2017
---


# Calendar.BaseCalendar Property (Project)

Gets a  **[Calendar](calendar-object-project.md)** object representing a base calendar. Read-only **Object**.


## Syntax

 _expression_. **BaseCalendar**

 _expression_ A variable that represents a **Calendar** object.


## Remarks

The  **BaseCalendar** property does not return any meaningful information for material resources. Setting a value returns a trappable error (error code 1101) when applied to material resources.


