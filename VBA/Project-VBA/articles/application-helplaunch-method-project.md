---
title: Application.HelpLaunch Method (Project)
keywords: vbapj.chm810
f1_keywords:
- vbapj.chm810
ms.prod: project-server
api_name:
- Project.Application.HelpLaunch
ms.assetid: 05e4e98c-bda7-5b41-372b-2f3752d2ab0e
ms.date: 06/08/2017
---


# Application.HelpLaunch Method (Project)

Starts a Help file.


## Syntax

 _expression_. **HelpLaunch**( ** _Filename_**, ** _ContextNumber_**, ** _Search_**, ** _SearchKey_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Filename_|Optional|**String**|The file name (with either .hlp or .chm extension) of the Help file to start. If FileName is not specified and Search is  **False**, the Project **Help** window appears with the navigation pane expanded.|
| _ContextNumber_|Optional|**Long**|The context number of a topic to display.|
| _Search_|Optional|**Boolean**|**True** if the **Help** window appears with the navigation pane expanded (CHM). If Search is **True**, ContextNumber is ignored. The default value is **False**.|
| _SearchKey_|Optional|**String**|Due to changes in the Project object model, this argument is no longer supported.|

### Return Value

 **Boolean**


