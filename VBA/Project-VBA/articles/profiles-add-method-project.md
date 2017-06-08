---
title: Profiles.Add Method (Project)
ms.prod: project-server
api_name:
- Project.Profiles.Add
ms.assetid: 056f912a-214f-8e23-338e-38e26b9d1e9d
ms.date: 06/08/2017
---


# Profiles.Add Method (Project)

Adds an account  **[Profile](profile-object-project.md)** object to the **Profiles** collection.


## Syntax

 _expression_. **Add**( ** _Name_**, ** _Server_**, ** _LoginType_**, ** _UserName_** )

 _expression_ A variable that represents a **Profiles** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**| The name of the profile.|
| _Server_|Required|**String**|A string representing the URL of the Project Server.|
| _LoginType_|Optional|**Long**|The login type for the Project Server. Can be one of the  **[PjLoginType](pjlogintype-enumeration-project.md)** constants. The default value is **pjWindowsLogin**.|
| _UserName_|Optional|**String**| A string representing the user name.|

### Return Value

 **Profile**


## Remarks

The UserName argument can be either a Project Server user name, if the LoginType is  **pjProjectServerLogin**, or a user name for a Windows account. For example, if the LoginType is **pjWindowsLogin**, a user name might be **DOMAIN\jsmith**.


## See also


#### Concepts


[Profiles Collection Object](profiles-object-project.md)
