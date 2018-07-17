---
title: Application.MailSystem Method (Project)
ms.prod: project-server
api_name:
- Project.Application.MailSystem
ms.assetid: 4ee9011c-f5f5-d0aa-0cd6-aa90130af4af
ms.date: 06/08/2017
---


# Application.MailSystem Method (Project)

Returns the type of e-mail system installed on the host machine.


## Syntax

 _expression_. **MailSystem**

 _expression_ A variable that represents an **Application** object.


### Return Value

[PjMailSystem](pjmailsystem-enumeration-project.md)


## Remarks

Can return one of the [PjMailSystem](pjmailsystem-enumeration-project.md) constants.


## Example

The following example sends the project file if the host machine is using MAPI.


```vb
Sub SendMAPI() 
 
 If Application.MailSystem = pjMAPI Then 
 MailSend To:="Jean Selva", Subject:="Sample Subject" 
 End If 
 
End Sub
```


