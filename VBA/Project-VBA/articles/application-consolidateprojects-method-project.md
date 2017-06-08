---
title: Application.ConsolidateProjects Method (Project)
keywords: vbapj.chm124
f1_keywords:
- vbapj.chm124
ms.prod: project-server
api_name:
- Project.Application.ConsolidateProjects
ms.assetid: 6f1f719c-09c0-076f-4680-24ac26a6538d
ms.date: 06/08/2017
---


# Application.ConsolidateProjects Method (Project)

Displays the data from one or more projects in a single window.


## Syntax

 _expression_. **ConsolidateProjects**( ** _Filenames_**, ** _NewWindow_**, ** _AttachToSources_**, ** _PoolResources_**, ** _HideSubtasks_**, ** _openPool_**, ** _UserID_**, ** _Password_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Filenames_|Optional|**String**|One or more file names of projects to consolidate.|
| _NewWindow_|Optional|**Boolean**|**True** if projects are inserted (consolidated) into a new project. **False** if projects are inserted into the active project at the selection point. The default value is **False**.|
| _AttachToSources_|Optional|**Boolean**|**True** if changes in the consolidated project affect source projects. The default value is **True**.|
| _PoolResources_|Optional|**Variant**|The  _PoolResources_ argument is ignored in Project. It is retained so that existing macros that use this argument do not cause errors.|
| _HideSubtasks_|Optional|**Boolean**|**True** if the subtasks of the projects specified with Filenames are hidden. The default value is **True**.|
| _openPool_|Optional|**Long**|The action to take when opening a resource pool or sharer file. When opening a master project, the value for this argument is also applied to the subprojects. Can be one of the following  **[PjPoolOpen](pjpoolopen-enumeration-project.md)** constants. The default value is **pjPromptPool**.|
| _UserID_|Optional|**Variant**| A user ID to use when accessing a project in a database. If Filenames does not refer to a database, **UserID** is ignored.|
| _Password_|Optional|**String**|A password to use when opening password-protected project files. If Password is incorrect or omitted and a file requires a password, the user is prompted for the password.|

### Return Value

 **Boolean**


## Remarks

To specify that a consolidated project should be inserted as read-only, append "(R/O)" to the file name in the  _Filenames_ argument.


## Example

The following example creates a consolidated project, prints a report, and closes the consolidated project without saving it.


```vb
Sub ConsolidatedReport() 
    ConsolidateProjects Filenames:="Project1.mpp" &; ListSeparator &; "Project2.mpp", NewWindow:=True 
    ReportPrint Name:="Critical Tasks" 
    FileClose Save:=pjDoNotSave 
End Sub
```


