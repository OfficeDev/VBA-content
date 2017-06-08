---
title: Application.DateOrder Property (Project)
ms.prod: project-server
api_name:
- Project.Application.DateOrder
ms.assetid: 9eba39c8-6e4a-3b8c-69c3-82e078269cda
ms.date: 06/08/2017
---


# Application.DateOrder Property (Project)

Gets the order of display of the day, month, and year in date values. Read-only  **PjDateOrder**.


## Syntax

 _expression_. **DateOrder**

 _expression_ A variable that represents an **Application** object.


## Remarks

The  **DateOrder** property can be one of the following **[PjDateOrder](pjdateorder-enumeration-project.md)** constants: **pjDayMonthYear**, **pjMonthDayYear**, or **pjYearMonthDay**.

Project sets the  **DateOrder** property equal to the corresponding value in the **Regional and Language Options** dialog box of the Microsoft Windows Control Panel. For example, if the current format is set to **French (France)**, the  **DateOrder** property value is 0 ( **pjDayMonthYear** ).


