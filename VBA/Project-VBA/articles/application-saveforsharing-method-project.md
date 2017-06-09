---
title: Application.SaveForSharing Method (Project)
keywords: vbapj.chm2133
f1_keywords:
- vbapj.chm2133
ms.prod: project-server
api_name:
- Project.Application.SaveForSharing
ms.assetid: a4f46990-aff1-52da-d1c7-7fd99e85d97a
ms.date: 06/08/2017
---


# Application.SaveForSharing Method (Project)

Saves a local copy of the active project for sharing, to make changes and then merge with the Project Server copy.


## Syntax

 _expression_. **SaveForSharing**( ** _Filename_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Filename_|Optional|**Variant**|Full path and the name of the project file saved for sharing.|

### Return Value

 **Boolean**


## Remarks

The  **SaveForSharing** method is available in Project Professional only. The original project on Project Server is marked as saved for sharing. The local copy of the project file can be modified and the changes merged with the original Project Server copy when you use the **Save As** command or the **FileSaveAs** method. If you try to reopen the Project Server copy before merging the local copy, Project Server disables sharing and prevents merging changes from the shared copy to the server.


