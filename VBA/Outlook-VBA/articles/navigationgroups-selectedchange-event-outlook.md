---
title: NavigationGroups.SelectedChange Event (Outlook)
keywords: vbaol11.chm2913
f1_keywords:
- vbaol11.chm2913
ms.prod: outlook
api_name:
- Outlook.NavigationGroups.SelectedChange
ms.assetid: eb55ed92-1925-9aaa-8fd6-9280cfc8aa47
ms.date: 06/08/2017
---


# NavigationGroups.SelectedChange Event (Outlook)

Occurs after the selection state is changed for a navigation folder contained in a  **Calendar** navigation module.


## Syntax

 _expression_ . **SelectedChange**( **_NavigationFolder_** )

 _expression_ A variable that represents a **NavigationGroups** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _NavigationFolder_|Required| **[NavigationFolder](navigationfolder-object-outlook.md)**|The selected navigation folder.|

## Remarks

This event occurs when the selection state changes for a folder in the  **Calendar** navigation module, either by a user checking or unchecking a folder in the **Calendar** navigation module of the Navigation Pane or by an add-in changing the value of the **[IsSelected](navigationfolder-isselected-property-outlook.md)** property for a **NavigationFolder** object contained in the **[NavigationGroups](navigationgroups-object-outlook.md)** collection of a **[CalendarModule](calendarmodule-object-outlook.md)** object.


## See also


#### Concepts


[NavigationGroups Object](navigationgroups-object-outlook.md)

