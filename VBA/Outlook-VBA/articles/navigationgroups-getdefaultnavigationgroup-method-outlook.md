---
title: NavigationGroups.GetDefaultNavigationGroup Method (Outlook)
keywords: vbaol11.chm2860
f1_keywords:
- vbaol11.chm2860
ms.prod: outlook
api_name:
- Outlook.NavigationGroups.GetDefaultNavigationGroup
ms.assetid: accdd554-1aa1-b254-7489-67673b889757
ms.date: 06/08/2017
---


# NavigationGroups.GetDefaultNavigationGroup Method (Outlook)

Returns the  **[NavigationGroup](navigationgroup-object-outlook.md)** that corresponds to the selected default shared folder group.


## Syntax

 _expression_ . **GetDefaultNavigationGroup**( **_DefaultFolderGroup_** )

 _expression_ A variable that represents a **NavigationGroups** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DefaultFolderGroup_|Required| **[OlGroupType](olgrouptype-enumeration-outlook.md)**|The type of navigation group to be retrieved.|

### Return Value

A  **NavigationGroup** object that represents the selected default folder group.


## Remarks

If the default navigation group specified in  _DefaultFolderGroup_ was deleted or otherwise doesn?t exist, it is automatically created if the parent **[NavigationModule](navigationmodule-object-outlook.md)** object supports the specified navigation group type. An error occurs if the parent **NavigationModule** object does not support creating this navigation group type.


## See also


#### Concepts


[NavigationGroups Object](navigationgroups-object-outlook.md)

