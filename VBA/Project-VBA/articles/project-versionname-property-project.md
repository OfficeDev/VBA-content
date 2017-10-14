---
title: Project.VersionName Property (Project)
keywords: vbapj.chm132790
f1_keywords:
- vbapj.chm132790
ms.prod: project-server
api_name:
- Project.Project.VersionName
ms.assetid: a1ad4584-39df-6897-c08d-d6cb94ee3cf4
ms.date: 06/08/2017
---


# Project.VersionName Property (Project)

Gets the version name of the project. Obsolete in Project. Read-only  **String**.


## Syntax

 _expression_. **VersionName**

 _expression_ A variable that represents a **Project** object.


## Remarks

In Project Server 2003, it is possible to have multiple projects with the same name but differentiated by version codes. In Office Project 2007 and later versions, each enterprise project must have a different name. The  **VersionName** property is an empty string ("").


