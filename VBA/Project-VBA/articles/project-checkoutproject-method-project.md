---
title: Project.CheckoutProject Method (Project)
keywords: vbapj.chm131078
f1_keywords:
- vbapj.chm131078
ms.prod: project-server
ms.assetid: 7b70a7c6-0f26-27b4-9a2d-b16f828864f3
ms.date: 06/08/2017
---


# Project.CheckoutProject Method (Project)
Checks out an open project that is currently in read-only mode.

## Syntax

 _expression_. **CheckoutProject**

 _expression_ A variable that represents a **Project** object.


### Return value

 **Nothing**


## Remarks

If the active project in Project Professional is in read-only mode, the  **CheckoutProject**.method checks out the project so that it is in read/write mode for editing. If the active project is already checked out, Project shows a dialog box with the message, "This project is already checked out to you on a different computer or Project Web App session."


## Example

The following example determines whether an open project is an enterprise project and is checked out. If the project is not checked out, the example tries to check out the project. If the project is already checked out to you, Project shows a dialog box with the error message, ''This project is already checked out to you on a different computer or Project Web App session." If the project is checked out by another user, Project shows a dialog box with the message, "To check out,  _DOMAIN\UserName_ must close the project in their session or contact your administrator to check in the project."


```vb
Sub CheckOutOpenEnterpriseProjects()
    Dim openProjects As Projects
    Dim proj As Project
    
    Set openProjects = Application.Projects
    
    On Error Resume Next
    
    For Each proj In openProjects
        If Application.IsCheckedOut(proj.Name) Then
            If proj.Type = pjProjectTypeEnterpriseCheckedOut Then
                Debug.Print "'" &; proj.Name &; "'" &; " is already checked out."
            ElseIf proj.Type = pjProjectTypeNonEnterprise Then
                Debug.Print "'" &; proj.Name &; "'" &; " is not an enterprise project."
            End If
        Else
            ' Check out the project whether it is active or not.
            proj.CheckoutProject
            Debug.Print "Attempted to check out: '" &; proj.Name &; "'"
        End If
    Next proj
End Sub
```


## See also


#### Concepts


[Project Object](project-object-project.md)
[Checkin Method](project-checkin-method-project.md)
#### Other resources


[Application.IsCheckedOut](application-ischeckedout-property-project.md)
[Application.ProjectCheckOut](application-projectcheckout-method-project.md)
