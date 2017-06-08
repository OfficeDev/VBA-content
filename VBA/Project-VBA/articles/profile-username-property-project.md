---
title: Profile.UserName Property (Project)
ms.prod: project-server
api_name:
- Project.Profile.UserName
ms.assetid: 8af2fe46-7218-39be-efd0-c7dd91f25ac7
ms.date: 06/08/2017
---


# Profile.UserName Property (Project)

Gets or sets the logon name of the current account profile. Read/write  **String**.


## Syntax

 _expression_. **UserName**

 _expression_ A variable that represents a **Profile** object.


## Remarks

The  **UserName** property of the **Profile** object shows the logon name. By comparison, the **[UserName](application-username-property-project.md)** property of the **Application** object shows the local user name.


## Example

If there are two account profiles, and the user named Jeff Smith logs on with the DOMAIN\jsmith account, the first statement in the following code shows  **DOMAIN\jsmith** in the **Immediate** pane of the VBE. The second statement shows **Jeff Smith**.


```vb
Debug.Print Profiles(2).UserName 
Debug.Print UserName
```


