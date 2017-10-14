---
title: Application.FileSaveOffline Method (Project)
keywords: vbapj.chm136
f1_keywords:
- vbapj.chm136
ms.prod: project-server
api_name:
- Project.Application.FileSaveOffline
ms.assetid: 109f95d5-be49-549f-fa39-3231207d61de
ms.date: 06/08/2017
---


# Application.FileSaveOffline Method (Project)

Saves the file for viewing offline.


## Syntax

 _expression_. **FileSaveOffline**( ** _Save_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Save_|Optional|**Long**|Can be one of the following  **PjSaveType** constants: **pjDoNotSave**, **pjSave**, or **pjPromptSave**. The default value is **pjPromptSave** for new project files and projects that have changed since the last save.|

### Return Value

 **Boolean**


