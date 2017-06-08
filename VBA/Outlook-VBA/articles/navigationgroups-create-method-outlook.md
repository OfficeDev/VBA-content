---
title: NavigationGroups.Create Method (Outlook)
keywords: vbaol11.chm2858
f1_keywords:
- vbaol11.chm2858
ms.prod: outlook
api_name:
- Outlook.NavigationGroups.Create
ms.assetid: 5f2bdcfc-4748-4170-7214-bcadc9e3dc36
ms.date: 06/08/2017
---


# NavigationGroups.Create Method (Outlook)

Creates and returns a new  **[NavigationGroup](navigationgroup-object-outlook.md)** object, added to the end of the **[NavigationGroups](navigationgroups-object-outlook.md)** collection.


## Syntax

 _expression_ . **Create**( **_GroupDisplayName_** )

 _expression_ A variable that represents a **NavigationGroups** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _GroupDisplayName_|Required| **String**|The value of the  **[Name](navigationgroup-name-property-outlook.md)** property for the new **NavigationGroup** object.|

### Return Value

A  **NavigationGroup** object that represents the new navigation group.


## Remarks

A  **NavigationGroups** collection can contain multiple **NavigationGroup** objects with the same **Name** property values.

An error occurs if an add-in attempts to add more than 50 navigation groups to a  **NavigationGroups** collection, or if an add-in attempts to add a **NavigationGroup** object to the **NavigationGroups** collection of a **[MailModule](mailmodule-object-outlook.md)** object.


## See also


#### Concepts


[NavigationGroups Object](navigationgroups-object-outlook.md)

