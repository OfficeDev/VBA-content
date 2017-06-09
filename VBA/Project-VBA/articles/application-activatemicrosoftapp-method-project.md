---
title: Application.ActivateMicrosoftApp Method (Project)
keywords: vbapj.chm131193
f1_keywords:
- vbapj.chm131193
ms.prod: project-server
api_name:
- Project.Application.ActivateMicrosoftApp
ms.assetid: a9b59db3-7ad2-8674-9026-090e161ef983
ms.date: 06/08/2017
---


# Application.ActivateMicrosoftApp Method (Project)

Activates a Microsoft application if the application is running or starts a new instance if the application is not running.


## Syntax

 _expression_. **ActivateMicrosoftApp**( ** _Index_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Long**|Specifies the Microsoft application to activate. Can be one of the  **[PjMSApplication](pjmsapplication-enumeration-project.md)** constants.|

