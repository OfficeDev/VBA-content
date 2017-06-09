---
title: Project.Activate Event (Project)
ms.prod: project-server
api_name:
- Project.Project.Activate
ms.assetid: fd3b89be-ea9a-5574-be1e-01e3d042a4a1
ms.date: 06/08/2017
---


# Project.Activate Event (Project)

Occurs when switching to the project from another project, including when the project is opened or created.


## Syntax

 _expression_. **Activate**( ** _pj_** )

 _expression_ A variable that represents a **Project** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _pj_|Required|**Project**|The project that was activated.|

## Remarks

When you switch between two windows showing the same project, the  **Activate** event for the project doesn't occur.

This event doesn't occur when you create a new window. 

Project events do not occur when the project is embedded in another document or application. 


## Example

The following example ensures that the project window is maximized whenever it is activated.


```vb
Private Sub Project_Activate(ByVal pj As MSProject.Project) 
    pj.Windows.ActiveWindow.WindowState = pjMaximized 
End Sub
```


