---
title: Application.OpenServerPage Method (Project)
keywords: vbapj.chm636
f1_keywords:
- vbapj.chm636
ms.prod: project-server
api_name:
- Project.Application.OpenServerPage
ms.assetid: 6b7e18fd-2ae1-47a0-45fb-58d6b6e27074
ms.date: 06/08/2017
---


# Application.OpenServerPage Method (Project)

Opens the specified page from Project Web App.


## Syntax

 _expression_. **OpenServerPage**( ** _Page_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Page_|Optional|**PjServerPage**|Specifies the page to open from Project Web App. Can be one of the  **[PjServerPage](pjserverpage-enumeration-project.md)** constants. The default is **pjServerPageApprovals**.|

### Return Value

 **Boolean**


## Remarks

Available in Project Professional only. Project must be connected to a Project Web App instance.


## Example

The following example opens the Issues page in the SharePoint workspace for the active project, and then opens the Project Center page in 

Project Web App

. Internet Explorer shows the pages in separate windows.




```vb
Sub OpenPages() 
    OpenServerPage Page:=pjServerPageIssues 
    OpenServerPage pjServerPageProjectCenter 
End Sub
```


