---
title: Project.MonthLabelDisplay Property (Project)
keywords: vbapj.chm132414
f1_keywords:
- vbapj.chm132414
ms.prod: project-server
api_name:
- Project.Project.MonthLabelDisplay
ms.assetid: ed6e783c-9f11-1ecf-7cf6-e8281a1892b2
ms.date: 06/08/2017
---


# Project.MonthLabelDisplay Property (Project)

Gets or sets the abbreviation for "month" that is displayed for values such as durations, delays, slack, and work. Read/write  **Integer**.


## Syntax

 _expression_. **MonthLabelDisplay**

 _expression_ A variable that represents a **Project** object.


## Remarks

The  **MonthLabelDisplay** property corresponds to the **Months** list on the **Advanced** tab of the **Project Options** dialog box. For example, setting the **MonthLabelDisplay** property to 1 sets the **Months** list to the second value in the list ("mon").

Values of the  **MonthLabelDisplay** property can be 0 to 2.


