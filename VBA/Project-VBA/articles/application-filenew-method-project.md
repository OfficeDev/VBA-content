---
title: Application.FileNew Method (Project)
keywords: vbapj.chm101
f1_keywords:
- vbapj.chm101
ms.prod: project-server
api_name:
- Project.Application.FileNew
ms.assetid: 59b5acd1-78dc-9fd2-d672-4cdd6a6005aa
ms.date: 06/08/2017
---


# Application.FileNew Method (Project)

Creates a new project.


## Syntax

 _expression_. **FileNew**( ** _SummaryInfo_**, ** _Template_**, ** _FileNewDialog_**, ** _FileNewWorkpane_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SummaryInfo_|Optional|**Boolean**|**True** if the ** Project Information** dialog box is displayed when creating the project. The default is equal to the corresponding setting on the **General** tab of the **Options** dialog box.|
| _Template_|Optional|**String**|A path and file name for a template to use when creating the project. If  **Template** is omitted, a blank project is created|
| _FileNewDialog_|Optional|**Boolean**|**True** if the **Templates** dialog box is displayed when creating the project. If **Template** is specified, **FileNewDialog**is ignored|
| _FileNewWorkpane_|Optional|**Boolean**|**True** if Project displays the ** New Project** workpane before creating a new file. The default value is **False**.|

### Return Value

 **Boolean**


