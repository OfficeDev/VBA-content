---
title: Application.CheckOut Method (Project)
keywords: vbapj.chm2332
f1_keywords:
- vbapj.chm2332
ms.prod: project-server
api_name:
- Project.Application.CheckOut
ms.assetid: 36e19455-a77d-46d5-c5c0-60f07feeba13
ms.date: 06/08/2017
---


# Application.CheckOut Method (Project)

Checks out the active project file if it is stored in a SharePoint library.


## Syntax

 _expression_. **CheckOut**

 _expression_ A variable that represents an **Application** object.


### Return Value

 **Boolean**


## Example

This example verifies that a project is not checked out by another user and can be checked out. If the project can be checked out, it copies the project to the local computer for editing.


```vb
Sub CheckOutProject(docCheckOut As String) 
 
 ' Determine if project can be checked out. 
 If Projects.CanCheckOut(docCheckOut) = True Then 
 Projects.CheckOut docCheckOut 
 Else 
 MsgBox "Unable to check out this project at this time." 
 End If 

```


