---
title: Project.DayLabelDisplay Property (Project)
ms.prod: project-server
api_name:
- Project.Project.DayLabelDisplay
ms.assetid: 6888b00a-3589-1e39-1394-c5089ec38521
ms.date: 06/08/2017
---


# Project.DayLabelDisplay Property (Project)

Gets or sets the abbreviation for "day" that is displayed for values such as durations, delays, slack, and work. Read/write  **Integer**.


## Syntax

 _expression_. **DayLabelDisplay**

 _expression_ A variable that represents a **Project** object.


## Remarks

The  **DayLabelDisplay** property corresponds to the **Days** list on the **Advanced** tab of the **Project Options** dialog box. For example, setting the **DayLabelDisplay** property to 1 sets the **Days** list to the second value in the list ("dy").

Values of the  **DayLabelDisplay** property can be 0 to 2.


