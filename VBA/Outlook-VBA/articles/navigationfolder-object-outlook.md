---
title: NavigationFolder Object (Outlook)
keywords: vbaol11.chm3201
f1_keywords:
- vbaol11.chm3201
ms.prod: outlook
api_name:
- Outlook.NavigationFolder
ms.assetid: c8d7aabb-58ba-df5e-ccdc-06f73db7726c
ms.date: 06/08/2017
---


# NavigationFolder Object (Outlook)

Represents a navigation folder displayed in a navigation group of a navigation module in the Navigation Pane.


## Remarks

Use the  **[Item](navigationfolders-item-method-outlook.md)** method to retrieve a **NavigationFolder** object from the **[NavigationFolders](navigationfolders-object-outlook.md)** collection of the parent **[NavigationGroup](navigationgroup-object-outlook.md)** object. Use the **[Add](navigationfolders-add-method-outlook.md)** method of the **NavigationFolders** collection to create a new **NavigationFolder** object based on an existing **[Folder](folder-object-outlook.md)** object.

Use the  **[Folder](navigationfolder-folder-property-outlook.md)** method to return or set the **Folder** object on which the **NavigationFolder** object is based.

Use the  **[IsSelected](navigationfolder-isselected-property-outlook.md)** property to determine if the navigation folder is selected and the **[Position](navigationfolder-position-property-outlook.md)** property to return or set the display position of the navigation folder within the Navigation Pane. You can also use the **[DisplayName](navigationfolder-displayname-property-outlook.md)** property to return the display name of the navigation folder within the Navigation Pane.

Use the  **[IsRemovable](navigationfolder-isremovable-property-outlook.md)** property to determine if a navigation folder can be removed from the **NavigationFolders** collection and the **[IsSideBySide](navigationfolder-issidebyside-property-outlook.md)** property to return or set the viewing mode for a navigation folder associated with a **[CalendarModule](calendarmodule-object-outlook.md)** object.


## Properties



|**Name**|
|:-----|
|[Application](navigationfolder-application-property-outlook.md)|
|[Class](navigationfolder-class-property-outlook.md)|
|[DisplayName](navigationfolder-displayname-property-outlook.md)|
|[Folder](navigationfolder-folder-property-outlook.md)|
|[IsRemovable](navigationfolder-isremovable-property-outlook.md)|
|[IsSelected](navigationfolder-isselected-property-outlook.md)|
|[IsSideBySide](navigationfolder-issidebyside-property-outlook.md)|
|[Parent](navigationfolder-parent-property-outlook.md)|
|[Position](navigationfolder-position-property-outlook.md)|
|[Session](navigationfolder-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
