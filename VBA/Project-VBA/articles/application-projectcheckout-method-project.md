---
title: Application.ProjectCheckOut Method (Project)
keywords: vbapj.chm2160
f1_keywords:
- vbapj.chm2160
ms.prod: project-server
ms.assetid: 4c6f065f-a853-8f42-e948-be7a76435c0b
ms.date: 06/08/2017
---


# Application.ProjectCheckOut Method (Project)
Checks out an open project if it is the active project.

## Syntax

 _expression_. **ProjectCheckOut** _(Name)_

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**String**|The name of the project|

### Return value

 **Boolean**


## Remarks

An open project must be active for the  **ProjectCheckOut** method to work. If the project is already checked out to you, Project shows a dialog box with the error message, ''This project is already checked out to you on a different computer or Project Web App session." If the project is checked out by another user, the error message is "To check out, _DOMAIN\UserName_ must close the project in their session or contact your administrator to check in the project."


## Example

The following example attempts to check out all projects that are opened as read-only.


```vb
Sub TestProjectCheckOut()
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
            ' Check out the project only if it is the active project.
            proj.Activate
            Application.ProjectCheckOut
            Debug.Print "Attempted to check out: '" &; proj.Name &; "'"
        End If
    Next proj
End Sub
```


## See also


#### Concepts


[Application Object](application-object-project.md)
[Project.Checkin Method](project-checkin-method-project.md)
#### Other resources


[IsCheckedOut Property](application-ischeckedout-property-project.md)
[Project.CheckoutProject Method](project-checkoutproject-method-project.md)
