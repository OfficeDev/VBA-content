---
title: Application.WindowNewWindow Method (Project)
keywords: vbapj.chm701
f1_keywords:
- vbapj.chm701
ms.prod: project-server
api_name:
- Project.Application.WindowNewWindow
ms.assetid: fe0c2bcb-7bee-3bec-9c47-3015938ae75d
ms.date: 06/08/2017
---


# Application.WindowNewWindow Method (Project)

Creates a window.


## Syntax

 _expression_. **WindowNewWindow**( ** _Projects_**, ** _View_**, ** _AllProjects_**, ** _ShowDialog_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Projects_|Optional|**String**|The names of one or more projects, separated by the list separator character. The new window contains data from the projects you specify. The default is to create a copy of the active window.|
| _View_|Optional|**String**|The name of an initial view for the new window. The default is equal to the value returned by the  **DefaultView** property.|
| _AllProjects_|Optional|**Boolean**|**True** if the new window contains data from all open projects. When **True**, AllProjects overrides Projects. The default value is **False**.|
| _ShowDialog_|Optional|**Boolean**|**True** if the **New Window** dialog box is displayed so that a view or project can be selected. The default value is **False.**|

### Return Value

 **Boolean**


## Example

The following example creates a window that combines the data from all open projects.


```vb
Sub NewCombineProjectsInNewWindow() 
 WindowNewWindow AllProjects:=True 
End Sub
```


