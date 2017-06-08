---
title: Application.EngagementInfo Method (Project)
keywords: vbapj.chm160
f1_keywords:
- vbapj.chm160
ms.assetid: 4e95d901-77a0-f1f7-b754-aefeb720e5ea
ms.date: 06/08/2017
ms.prod: project-server
---


# Application.EngagementInfo Method (Project)

Displays the engagement information dialog box user interface for the  **Resource Plan** view. Introduced in Office 2016.


## Syntax

 _expression_. **EngagementInfo**( _EngagementUniqueID_,  _EngagementUniqueID_,  _ResourceUniqueID_,  _Description_,  _StartDate_,  _FinishDate_,  _Units_,  _Work_,  _ShowDialog_)

 _expression_ A variable that represents a **Application** object.


### Parameters


|||||
|:-----|:-----|:-----|:-----|
|**Name**|**Required/Optional**|**Value**|**Description**|
| _EngagementUniqueID_|Optional|Dword|The unique ID of the engagement.|
| _ResourceUniqueID_|Optional|Dword|The unique ID of the resource.|
| _Description_|Optional|String|A description of the engagement.|
| _StartDate_|Optional|Date|The earliest date the resource can work on the engagement.|
| _FinishDate_|Optional|Date|The latest date the resource can work on the engagement.|
| _Units_|Optional|Real|The assignment unit the resource can work on the engagement.|
| _Work_|Optional|Real|The amount of work requested or approved for the engagement.|
| _ShowDialog_|Required|Boolean|Default=1; Displayed|

### Return Value

 **BOOLEAN**


## See also


#### Concepts


[Application Object (Project)](application-object-project.md)

