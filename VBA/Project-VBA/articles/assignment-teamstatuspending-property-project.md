---
title: Assignment.TeamStatusPending Property (Project)
ms.prod: project-server
api_name:
- Project.Assignment.TeamStatusPending
ms.assetid: 8e403925-225e-a1e9-121c-6f9353578150
ms.date: 06/08/2017
---


# Assignment.TeamStatusPending Property (Project)

 **True** if a response has not been received for a progress request message. **False** if the assigned resource has sent a response. Read/write **Boolean**.


## Syntax

 _expression_. **TeamStatusPending**

 _expression_ A variable that represents an **Assignment** object.


## Remarks

To see whether the team member for the assignment has responded to an Update Progress request, add the  **TeamStatusPending** field to the **Task Usage** or **Resource Usage** view.


