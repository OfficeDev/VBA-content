---
title: Application.CompareProjectVersions Method (Project)
keywords: vbapj.chm2183
f1_keywords:
- vbapj.chm2183
ms.prod: project-server
api_name:
- Project.Application.CompareProjectVersions
ms.assetid: 82af9450-0cec-f7b4-df5c-81ecea3b662f
ms.date: 06/08/2017
---


# Application.CompareProjectVersions Method (Project)

Displays the  **Compare Project Versions** dialog box to compare two versions of a project.


## Syntax

 _expression_. **CompareProjectVersions**

 _expression_ An expression that returns an **Application** object.


### Return Value

 **Boolean**


## Remarks

The  **CompareProjectVersions** method is equivalent to the **Compare Projects** command in the **Reports** group of the **Project** tab on the Ribbon. If you want programmatic control of the project comparison feature (such as whether difference columns are displayed), use the **[CreateComparisonReport](application-createcomparisonreport-method-project.md)** method.


## Example

The following example checks whether a project is open before calling the  **CompareProjectVersions** method. If a project is open, the code checks whether there are either tasks or resources in the project before calling the method.


```vb
Sub CompareVersions () 
    If Projects.Count = 0 Then 
        MsgBox "You must have at least one project open before you can compare projects." 
    Exit Sub 
 
    ElseIf ActiveProject.Tasks.Count = 0 Then 
        If ActiveProject.ResourceCount = 0 Then 
            MsgBox "There are no task or resources in the current project." &; vbCrLf &; _ 
                "Open a project with either tasks or resources before creating a comparison report.", _ 
                vbInformation 
            Exit Sub 
        End If 
    End If 
 
    CompareProjectVersions 
End Sub
```


