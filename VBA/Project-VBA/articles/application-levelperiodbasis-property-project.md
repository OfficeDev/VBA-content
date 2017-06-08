---
title: Application.LevelPeriodBasis Property (Project)
ms.prod: project-server
api_name:
- Project.Application.LevelPeriodBasis
ms.assetid: 24a13a72-8a3d-e59b-d912-6847f79019e1
ms.date: 06/08/2017
---


# Application.LevelPeriodBasis Property (Project)

Gets or sets the period by which resources are checked for overallocations. Read/write  **PjLevelPeriodBasis**.


## Syntax

 _expression_. **LevelPeriodBasis**

 _expression_ A variable that represents an **Application** object.


## Remarks

The  **LevelPeriodBasis** property can be one of the following **[PjLevelPeriodBasis](pjlevelperiodbasis-enumeration-project.md)** constants: **pjMinuteByMinute**, **pjHourByHour**, **pjDayByDay**, **pjWeekByWeek**, or **pjMonthByMonth**.

You can also set the  **LevelPeriodBasis** property in the **Resource Leveling** dialog box. To access the setting, click **Leveling Options** on the **Resource** tab of the Ribbon, and then set the overallocation leveling period basis in the drop-down list in the **Leveling calculations** section.


