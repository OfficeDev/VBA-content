---
title: Projects.Add Method (Project)
ms.prod: project-server
api_name:
- Project.Projects.Add
ms.assetid: 51629c33-1521-bfee-edf7-bed792d393c1
ms.date: 06/08/2017
---


# Projects.Add Method (Project)

Adds a  **Project** object to a **Projects** collection.


## Syntax

 _expression_. **Add**( ** _DisplayProjectInfo_**, ** _Template_**, ** _FileNewDialog_** )

 _expression_ A variable that represents a **Projects** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DisplayProjectInfo_|Optional|**Boolean**|**True** if the ** Project Information** dialog box appears when a new project is created. The default value is **False**.|
| _Template_|Optional|**String**|A path and file name for a template to use when creating the project. If Template is omitted, a blank project is created.|
| _FileNewDialog_|Optional|**Boolean**|**True** if the **Templates** dialog box is displayed when creating the project. If Template is specified, FileNewDialog is ignored.|

### Return Value

 **Project**


## See also


#### Concepts


[Projects Collection Object](projects-object-project.md)
