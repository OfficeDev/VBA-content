---
title: Projects.CheckOut Method (Project)
keywords: vbapj.chm132593
f1_keywords:
- vbapj.chm132593
ms.prod: project-server
api_name:
- Project.Projects.CheckOut
ms.assetid: 2de8fef7-150b-4f67-4677-507f5d2a258f
ms.date: 06/08/2017
---


# Projects.CheckOut Method (Project)

Checks out the specified file if it is stored in a SharePoint document library.


## Syntax

 _expression_. **CheckOut**( ** _Filename_** )

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
  
 ' Determine whether project can be checked out.  
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
