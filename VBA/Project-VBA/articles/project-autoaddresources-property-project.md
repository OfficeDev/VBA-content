---
title: Project.AutoAddResources Property (Project)
ms.prod: project-server
api_name:
- Project.Project.AutoAddResources
ms.assetid: b8e048f8-5bc1-740f-d397-e6f0ddf77a3a
ms.date: 06/08/2017
---


# Project.AutoAddResources Property (Project)

 **True** if new resources are automatically created as they are assigned. **False** if Project prompts before creating new resources. Read/write **Boolean**.


## Syntax

 _expression_. **AutoAddResources**

 _expression_ A variable that represents a **Project** object.


## Example

The following example prompts the user to set the  **AutoAddResources**, **AutoCalculate**, **AutoLinkTasks**, **AutoSplitTasks**, and **AutoTrack** properties.


```vb
Sub PromptForAutoPropertySettings() 
    Dim I As Integer ' Used in For...Next loop 
    Dim Prompts(5) As String ' Prompts to display on the screen 
    Dim Response As Long ' User response to prompt 
    Dim Responses(5) As Long ' Used to store user responses 
 
    ' Set each prompt. 
    Prompts(1) = "Automatically create new resources as they are assigned?" 
    Prompts(2) = "Automatically recalculate a project when a value, such as a date or cost, changes?" 
    Prompts(3) = "Automatically link sequential tasks when you cut, move, or insert tasks?" 
    Prompts(4) = "Automatically split tasks into parts for work complete and work remaining?" 
    Prompts(5) = "Automatically update the remaining work and cost for a resource when the completion percentage of one of the resource's tasks changes?" 
 
    ' Display each prompt, and store the user's responses. 
    For I = 1 To 5 
        Response = MsgBox(Prompts(I), vbYesNo) 
        Responses(I) = (Response = vbYes) 
    Next I 
 
    ' Set the automatic properties according to the user's responses. 
    ActiveProject.AutoAddResources = Responses(1) 
    Calculation = Responses(2) 
    ActiveProject.AutoLinkTasks = Responses(3) 
    ActiveProject.AutoSplitTasks = Responses(4) 
    ActiveProject.AutoTrack = Responses(5) 
End Sub
```


