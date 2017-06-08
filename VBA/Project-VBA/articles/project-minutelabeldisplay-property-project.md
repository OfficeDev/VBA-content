---
title: Project.MinuteLabelDisplay Property (Project)
keywords: vbapj.chm132410
f1_keywords:
- vbapj.chm132410
ms.prod: project-server
api_name:
- Project.Project.MinuteLabelDisplay
ms.assetid: 7cf43dda-ae9b-ed06-027e-740ba855e7f1
ms.date: 06/08/2017
---


# Project.MinuteLabelDisplay Property (Project)

Gets or sets the abbreviation for "minute" that is displayed for values such as durations, delays, slack, and work. Read/write  **Integer**.


## Syntax

 _expression_. **MinuteLabelDisplay**

 _expression_ A variable that represents a **Project** object.


## Remarks

The  **MinuteLabelDisplay** property corresponds to the **Minutes** list on the **Advanced** tab of the **Project Options** dialog box. For example, setting the **MinuteLabelDisplay** property to 1 sets the **Minutes** list to the second value in the list ("min").

Values of the  **MinuteLabelDisplay** property can be 0 to 2.


