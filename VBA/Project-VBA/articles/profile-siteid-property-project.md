---
title: Profile.SiteId Property (Project)
ms.prod: project-server
ms.assetid: 18d72450-e7d6-55b7-733c-45db023469c5
ms.date: 06/08/2017
---


# Profile.SiteId Property (Project)
Gets the GUID of the Project Web App instance for the active profile. Read-only  **String**.

## Syntax

 _expression_. **SiteId**

 _expression_ A variable that represents a **Profile** object.


## Remarks

If the active profile is for the local computer, the  **SiteId** property is an empty string.


## Example

If you enter the following statement in the Immediate pane of the VBE, the statement returns the GUID of the connected Project Web App instance, for example,  `{37522002-393E-4594-8017-9068DB816220}`.


```vb
? Profiles.ActiveProfile.SiteId
```


## Property value

 **STRING**


## See also


#### Concepts


[Profile Object](profile-object-project.md)
