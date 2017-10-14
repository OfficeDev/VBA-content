---
title: Projects.CanCheckOut Method (Project)
keywords: vbapj.chm132591
f1_keywords:
- vbapj.chm132591
ms.prod: project-server
api_name:
- Project.Projects.CanCheckOut
ms.assetid: 330f28a3-d785-ae5d-0f64-8e02ac52d8d6
ms.date: 06/08/2017
---


# Projects.CanCheckOut Method (Project)

Indicates whether Project can check out the specified project from a SharePoint document library.


## Syntax

 _expression_. **CanCheckOut**( ** _Filename_** )

 _expression_ A variable that represents a **Projects** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Filename_|Required|**String**|The name of the file to check out.|

### Return Value

 **Boolean**


## Example

The following example verifies that a project is not checked out by another user. If the project can be checked out, the example copies the project to the local computer for editing.


```vb
Sub CheckOutProject(docCheckOut As String)  
    ' Determine if project can be checked out.  
    If Projects.CanCheckOut(docCheckOut) = True Then  
        Projects.CheckOut docCheckOut  
    Else  
        MsgBox "Unable to check out this project at this time."  
    End If  
End Sub
```


## See also


#### Concepts


[Projects Collection Object](projects-object-project.md)
