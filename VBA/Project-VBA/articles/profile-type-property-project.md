---
title: Profile.Type Property (Project)
ms.prod: project-server
api_name:
- Project.Profile.Type
ms.assetid: ff5c3939-cfa6-c098-5fc4-180a4573ecb0
ms.date: 06/08/2017
---


# Profile.Type Property (Project)

 Gets a value that specifies whether the account profile being used is a local profile or for Project Server. Read-only **PjProfileType**.


## Syntax

 _expression_. **Type**

 _expression_ A variable that represents a **Profile** object.


## Remarks

The Type property can be one of the following  **[PjProfileType](pjprofiletype-enumeration-project.md)** constants: **pjLocalProfile** or **pjServerProfile**.

The  **Project Server Accounts** dialog box shows the number and order of profiles. Use `Profiles.Count` to programmatically determine the number of account profiles defined in Project Professional.


## Example

If the second account profile is a Project Server account, the following statement returns 1, which is the value of the  **pjServerProfile** constant.


```vb
Debug.Print Profiles(2).Type
```


