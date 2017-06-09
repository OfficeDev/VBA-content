---
title: Profile Object (Project)
ms.prod: project-server
api_name:
- Project.Profile
ms.assetid: 92ae9d1a-ea4d-1814-1655-f0798f4b18d0
ms.date: 06/08/2017
---


# Profile Object (Project)


 

Represents an account profile in Project Professional. The  **Profile** object is a member of the **[Profiles](profiles-object-project.md)** collection.
 
If the second account profile is a Project Server account, the following statement returns the value 1, which is the value of the  **pjServerProfile** constant in the **[PjProfileType](pjprofiletype-enumeration-project.md)** enumeration.
 



```
Debug.Print Profiles(2).Type
```


## Remarks

The  **Project Server Accounts** dialog box shows the number and order of profiles. Use `Profiles.Count` to programmatically determine the number of account profiles defined in Project Professional.
 

 

## Methods



|**Name**|
|:-----|
|[Delete](profile-delete-method-project.md)|

## Properties



|**Name**|
|:-----|
|[ConnectionState](profile-connectionstate-property-project.md)|
|[LoginType](profile-logintype-property-project.md)|
|[Name](profile-name-property-project.md)|
|[Server](profile-server-property-project.md)|
|[SiteId](profile-siteid-property-project.md)|
|[Type](profile-type-property-project.md)|
|[UserName](profile-username-property-project.md)|

