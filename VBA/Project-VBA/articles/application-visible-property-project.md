---
title: Application.Visible Property (Project)
ms.prod: project-server
api_name:
- Project.Application.Visible
ms.assetid: 43bf25de-4908-1fad-e5d5-9fba21e8b03c
ms.date: 06/08/2017
---


# Application.Visible Property (Project)

 **True** if the application is visible. Read/write **Boolean**.


## Syntax

 _expression_. **Visible**

 _expression_ A variable that represents an **Application** object.


## Remarks

The  **Visible** property can only be set to **False** if the **Application**. **[UserControl](application-usercontrol-property-project.md)** property is **False** and there are no visible projects. If the **UserControl** property is **True**, the Project application is under user control rather than programmatic control, and the **Visible** property is also **True**.


