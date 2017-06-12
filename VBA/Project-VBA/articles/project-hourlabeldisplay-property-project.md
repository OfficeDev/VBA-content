---
title: Project.HourLabelDisplay Property (Project)
keywords: vbapj.chm132338
f1_keywords:
- vbapj.chm132338
ms.prod: project-server
api_name:
- Project.Project.HourLabelDisplay
ms.assetid: 6dc5f65b-d509-5d4a-a550-52c92b43534e
ms.date: 06/08/2017
---


# Project.HourLabelDisplay Property (Project)

Gets or sets the abbreviation for "hour" that is displayed for values such as durations, delays, slack, and work. Read/write  **Integer**.


## Syntax

 _expression_. **HourLabelDisplay**

 _expression_ A variable that represents a **Project** object.


## Remarks

The  **HourLabelDisplay** property corresponds to the **Hours** list on the **Advanced** tab of the **Project Options** dialog box. For example, setting the **HourLabelDisplay** property to 1 sets the **Hours** list to the second value in the list ("hr").

Values of the  **HourLabelDisplay** property can be 0 to 2.


