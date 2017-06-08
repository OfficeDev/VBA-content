---
title: Application.ObjectInsert Method (Project)
keywords: vbapj.chm221
f1_keywords:
- vbapj.chm221
ms.prod: project-server
api_name:
- Project.Application.ObjectInsert
ms.assetid: 2956dd32-9e28-76e9-c991-12650ee48576
ms.date: 06/08/2017
---


# Application.ObjectInsert Method (Project)

Displays the  **Insert Object** dialog box, which prompts the user to insert an object.


## Syntax

 _expression_. **ObjectInsert**

 _expression_ A variable that represents an **Application** object.


### Return Value

 **Boolean**


## Remarks

The  **ObjectInsert** method is equivalent to the **Object** command. For an example of how to use the **Object** command, see the **[ObjectChangeIcon](application-objectchangeicon-method-project.md)** method.

The  **ObjectInsert** method has no effect if the active view is a combination view, Calendar view, Network Diagram, Relationship Diagram, or Resource Graph. In addition to these views, the **ObjectInsert** method has no effect unless a non-null task or resource is selected in the Task or Resource Sheet views.


