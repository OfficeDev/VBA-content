---
title: Application.AMText Property (Project)
keywords: vbapj.chm131364
f1_keywords:
- vbapj.chm131364
ms.prod: project-server
api_name:
- Project.Application.AMText
ms.assetid: 92a8d781-79ac-ebfa-8419-31cbd140e505
ms.date: 06/08/2017
---


# Application.AMText Property (Project)

Gets the text that Project displays next to morning hours in the 12-hour time format. Read/write  **String**.


## Syntax

 _expression_. **AMText**

 _expression_ A variable that represents an **Application** object.


## Remarks

Project sets the  **AMText** and **PMText** properties equal to the corresponding values in the **Regional and Language Options** dialog box opened from the Microsoft Windows Control Panel.


 **Note**  Although the VBA Object Browser shows  **AMText** as read-write, you cannot set the value using the **PMText** property.


