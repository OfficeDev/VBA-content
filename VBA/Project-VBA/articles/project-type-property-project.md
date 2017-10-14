---
title: Project.Type Property (Project)
ms.prod: project-server
api_name:
- Project.Project.Type
ms.assetid: 13393b8e-283d-d816-283e-f363b83eac91
ms.date: 06/08/2017
---


# Project.Type Property (Project)

Gets the type of a project. Read-only  **PjProjectType**.


## Syntax

 _expression_. **Type**

 _expression_ A variable that represents a **Project** object.


## Remarks

The  **Type** property can be one of the **[PjProjectType](pjprojecttype-enumeration-project.md)** constants.


## Example

The following example determines whether an open project is an enterprise project and is checked out. If the project is not checked out, the example tries to check out the project. If the project is checked out by another user, Project shows a dialog box with the message, "To check out, DOMAIN\UserName must close the project in their session or contact your administrator to check in the project."


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
            proj.CheckoutProject
            Debug.Print "Attempted to check out: '" &; proj.Name &; "'"
        End If
    Next proj
End Sub
```


## See also


#### Concepts


[Project Object](project-object-project.md)
[PjProjectType Enumeration](pjprojecttype-enumeration-project.md)
#### Other resources


[CheckoutProject Method](project-checkoutproject-method-project.md)
[Application.IsCheckedOut Property](application-ischeckedout-property-project.md)
