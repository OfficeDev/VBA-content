---
title: Application.UserName Property (Project)
ms.prod: project-server
api_name:
- Project.Application.UserName
ms.assetid: c501ef16-f4c8-3c08-69b8-3e9756db8336
ms.date: 06/08/2017
---


# Application.UserName Property (Project)

Gets or sets the local name of the current user. Read/write  **String**.


## Syntax

 _expression_. **UserName**

 _expression_ A variable that represents an **Application** object.


## Remarks

 The **UserName** property of the **Application** object shows the local user name. By comparison, the **[UserName](profile-username-property-project.md)** property of the **Profile** object shows the logon name.

Use the  **UserName** property to customize Project options or macros for a particular user. For example, suppose you have written a macro named **PrintReport** that prints the Mine.mpp report when you press CTRL+R, but another user wants to use the same shortcut keys to print the Yours.mpp report. You can edit the **PrintReport** macro so that it checks the **UserName** property and then prints Mine.mpp if you are the current user or prints Yours.mpp if you are not the current user.


 **Note**  The  **UserName** property is the local name but can be changed to a different value. The **Author** field in the **Project Properties** dialog box is the logon name of the user by default.


## Example

The following example sets preferences according to the name of the current user.


```vb
Sub GetUserName() 
 
    ' Get the user name. 
    UserName = InputBox$("What's your name?", , UserName) 
 
    ' If user is Jeff Smith, then set certain preferences. 
    If UserName = "Jeff Smith" Then 
        DisplayScheduleMessages = False 
        BarRounding On:=False 
        Calculation = True 
    ' Otherwise, set default preferences. 
    Else 
        DisplayScheduleMessages = True 
        BarRounding On:=True 
        Calculation = False 
    End If
End Sub
```


