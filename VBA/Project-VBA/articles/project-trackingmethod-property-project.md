---
title: Project.TrackingMethod Property (Project)
ms.prod: project-server
api_name:
- Project.Project.TrackingMethod
ms.assetid: cda3f127-5fad-f486-f02d-6d6eeb0d5588
ms.date: 06/08/2017
---


# Project.TrackingMethod Property (Project)

Gets or sets the tracking method used by Project Server for the project. Read/write  **PjProjectServerTrackingMethod**.


## Syntax

 _expression_. **TrackingMethod**

 _expression_ A variable that represents a **Project** object.


## Remarks

The  **TrackingMethod** property is available only in Project Professional, when the project is opened from Project Server. It can be one of the following **[PjProjectServerTrackingMethod](pjprojectservertrackingmethod-enumeration-project.md)** constants: **pjTrackingMethodDefault**, **pjTrackingMethodPercentComplete**, **pjTrackingMethodSpecifyHours**, or **pjTrackingMethodTotalAndRemaining**.


