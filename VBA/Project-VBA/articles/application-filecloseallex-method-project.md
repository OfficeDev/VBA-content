---
title: Application.FileCloseAllEx Method (Project)
keywords: vbapj.chm104
f1_keywords:
- vbapj.chm104
ms.prod: project-server
api_name:
- Project.Application.FileCloseAllEx
ms.assetid: 95c7c89f-cfb0-f881-a31b-70ae951fb3f1
ms.date: 06/08/2017
---


# Application.FileCloseAllEx Method (Project)

Closes all projects.


## Syntax

 _expression_. **FileCloseAllEx**( ** _Save_**, ** _CheckIn_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Save_|Optional|**Long**|Can be one of the following  **PjSave** constants: **pjDoNotSave**, **pjSave**, or **pjPromptSave**. The default value is **pjPromptSave** for new project files and projects that have changed since the last save.|
| _CheckIn_|Optional|**Variant**|**True** if files are checked in after closing.|

### Return Value

 **Boolean**


