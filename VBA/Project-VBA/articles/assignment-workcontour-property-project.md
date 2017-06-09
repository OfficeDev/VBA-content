---
title: Assignment.WorkContour Property (Project)
keywords: vbapj.chm132828
f1_keywords:
- vbapj.chm132828
ms.prod: project-server
api_name:
- Project.Assignment.WorkContour
ms.assetid: a47a3012-7e5e-febb-d023-368c7c01e065
ms.date: 06/08/2017
---


# Assignment.WorkContour Property (Project)

Gets or sets the type of work contour for the assignment. Read/write  **PjWorkContourType**.


## Syntax

 _expression_. **WorkContour**

 _expression_ A variable that represents an **Assignment** object.


## Remarks

The  **WorkContour** property can be one of the following **[PjWorkContourType](pjworkcontourtype-enumeration-project.md)** constants: **pjBackLoaded**, **pjBell**, **pjContour**, **pjDoublePeak**, **pjEarlyPeak**, **pjFlat**, **pjFrontLoaded**, **pjLatePeak**, or **pjTurtle**. The default value is **pjFlat**.


