---
title: NavigationGroups Object (Outlook)
keywords: vbaol11.chm3022
f1_keywords:
- vbaol11.chm3022
ms.prod: outlook
api_name:
- Outlook.NavigationGroups
ms.assetid: 07206203-36a9-7467-3a89-24fa2a7c2b1f
ms.date: 06/08/2017
---


# NavigationGroups Object (Outlook)

Contains a set of  **[NavigationGroup](navigationgroup-object-outlook.md)** objects that represent the navigation groups displayed by a navigation module in the Navigation Pane.


## Remarks

Use the  **[NavigationGroups](mailmodule-navigationgroups-property-outlook.md)** property of the parent navigation module, such as a **[MailModule](mailmodule-object-outlook.md)** object, to return a **NavigationGroups** object.

Use the  **[Create](navigationgroups-create-method-outlook.md)** method to create a new **NavigationGroup** object and add it to the collection. Use the **[Item](navigationgroups-item-method-outlook.md)** method to retrieve a **NavigationGroup** object from the collection. Use the **[Delete](navigationgroups-delete-method-outlook.md)** method of the **NavigationGroups** collection to create a new **NavigationGroup** object.

Use the  **[GetDefaultNavigationGroup](navigationgroups-getdefaultnavigationgroup-method-outlook.md)** to return the default navigation group for a specified group type.


## Events



|**Name**|
|:-----|
|[NavigationFolderAdd](navigationgroups-navigationfolderadd-event-outlook.md)|
|[NavigationFolderRemove](navigationgroups-navigationfolderremove-event-outlook.md)|
|[SelectedChange](navigationgroups-selectedchange-event-outlook.md)|

## Methods



|**Name**|
|:-----|
|[Create](navigationgroups-create-method-outlook.md)|
|[Delete](navigationgroups-delete-method-outlook.md)|
|[GetDefaultNavigationGroup](navigationgroups-getdefaultnavigationgroup-method-outlook.md)|
|[Item](navigationgroups-item-method-outlook.md)|

## Properties



|**Name**|
|:-----|
|[Application](navigationgroups-application-property-outlook.md)|
|[Class](navigationgroups-class-property-outlook.md)|
|[Count](navigationgroups-count-property-outlook.md)|
|[Parent](navigationgroups-parent-property-outlook.md)|
|[Session](navigationgroups-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
