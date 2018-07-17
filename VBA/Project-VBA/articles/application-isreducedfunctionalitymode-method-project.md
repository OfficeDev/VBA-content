---
title: Application.IsReducedFunctionalityMode Method (Project)
ms.prod: project-server
api_name:
- Project.Application.IsReducedFunctionalityMode
ms.assetid: d53320db-377d-2e78-10b2-03af8d8bded3
ms.date: 06/08/2017
---


# Application.IsReducedFunctionalityMode Method (Project)

Indicates whether the installed Project application is in reduced functionality mode.


## Syntax

 _expression_. **IsReducedFunctionalityMode**

 _expression_ A variable that represents an **Application** object.


### Return Value

 **Boolean**


## Remarks

If a user does not activate Project after installing it, after a period of time the application will start only in reduced-functionality mode. In this mode, the user cannot save changes to projects or create a new project. To be able to use the full functionality, the user must activate the Project application.


