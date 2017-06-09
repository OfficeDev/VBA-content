---
title: Application.LinksBetweenProjects Method (Project)
keywords: vbapj.chm245
f1_keywords:
- vbapj.chm245
ms.prod: project-server
api_name:
- Project.Application.LinksBetweenProjects
ms.assetid: 63962df8-05ef-f3b4-7ad7-4c75b50ac398
ms.date: 06/08/2017
---


# Application.LinksBetweenProjects Method (Project)

Specifies whether the  **Links between Projects** dialog box appears when opening a project containing cross-project links.


## Syntax

 _expression_. **LinksBetweenProjects**( ** _AcceptAll_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _AcceptAll_|Optional|**Boolean**|**True** if all changes to external predecessors and successors are accepted. **False** if the **Links between Projects** dialog box appears. The default value is **False**.|

### Return Value

 **Boolean**


