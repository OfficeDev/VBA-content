---
title: NavigationGroups.Delete Method (Outlook)
keywords: vbaol11.chm2859
f1_keywords:
- vbaol11.chm2859
ms.prod: outlook
api_name:
- Outlook.NavigationGroups.Delete
ms.assetid: b5bb08c4-9cf1-4ed7-9522-0096f1016e5b
ms.date: 06/08/2017
---


# NavigationGroups.Delete Method (Outlook)

Deletes the specified  **[NavigationGroup](navigationgroup-object-outlook.md)** object from the **[NavigationGroups](navigationgroups-object-outlook.md)** collection.


## Syntax

 _expression_ . **Delete**( **_Group_** )

 _expression_ A variable that represents a **NavigationGroups** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Group_|Required| **NavigationGroup**|The navigation group to be deleted.|

## Remarks

The  **Delete** method raises an error if:


-  The navigation group specified in _Group_ contains navigation folders in its **[NavigationFolders](navigationfolders-object-outlook.md)** collection.
    
- The  **[GroupType](navigationgroup-grouptype-property-outlook.md)** property of the navigation group specified in _Group_ is set to **olMyFoldersGroup** .
    
- The parent of the  **NavigationGroups** collection is a **[MailModule](mailmodule-object-outlook.md)** object.
    

## See also


#### Concepts


[NavigationGroups Object](navigationgroups-object-outlook.md)

