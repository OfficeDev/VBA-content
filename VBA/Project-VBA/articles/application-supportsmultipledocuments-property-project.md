---
title: Application.SupportsMultipleDocuments Property (Project)
keywords: vbapj.chm132676
f1_keywords:
- vbapj.chm132676
ms.prod: project-server
api_name:
- Project.Application.SupportsMultipleDocuments
ms.assetid: d5f1daf1-21b0-3c6c-44b2-8e3f665c7055
ms.date: 06/08/2017
---


# Application.SupportsMultipleDocuments Property (Project)

Always  **True** for Project and any other application that supports multiple documents (projects). Read-only **Boolean**.


## Syntax

 _expression_. **SupportsMultipleDocuments**

 _expression_ A variable that represents an **Application** object.


## Remarks

The  **SupportsMultipleDocuments** property is useful with Automation. For example, suppose you want to open a second document in the application referred to by a variable. If the variable refers to one of several possible applications, you may want to use the **SupportsMultipleDocuments** property to confirm that the application currently referenced by the variable can have more than one document open at a time.


