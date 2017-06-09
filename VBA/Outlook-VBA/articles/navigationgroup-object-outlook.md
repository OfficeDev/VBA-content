---
title: NavigationGroup Object (Outlook)
keywords: vbaol11.chm3199
f1_keywords:
- vbaol11.chm3199
ms.prod: outlook
api_name:
- Outlook.NavigationGroup
ms.assetid: a96eb2b1-af1f-71b2-6a0b-dcb5078beb1f
ms.date: 06/08/2017
---


# NavigationGroup Object (Outlook)

Represents a navigation group displayed by a navigation module in the Navigation Pane.


## Remarks

Use the  **[Item](navigationgroups-item-method-outlook.md)** method to retrieve a **NavigationGroup** object from the **[NavigationGroups](navigationgroups-object-outlook.md)** collection of a parent navigation module, such as a **[MailModule](mailmodule-object-outlook.md)** object. Use the **[Create](navigationgroups-create-method-outlook.md)** method of the **NavigationGroups** collection to create a new **NavigationGroup** object.

Use the  **[GroupType](navigationgroup-grouptype-property-outlook.md)** property to determine the group type of the navigation group and the **[Position](navigationgroup-position-property-outlook.md)** property to return or set the display position of the navigation group within the Navigation Pane. You can also use the **[Name](navigationgroup-name-property-outlook.md)** property to return or set the display name of the navigation group within the Navigation Pane.

Use the  **[NavigationFolders](navigationgroup-navigationfolders-property-outlook.md)** property to return a **[NavigationFolders](navigationfolders-object-outlook.md)** object containing the navigation folders for the specified navigation group.


## Properties



|**Name**|
|:-----|
|[Application](navigationgroup-application-property-outlook.md)|
|[Class](navigationgroup-class-property-outlook.md)|
|[GroupType](navigationgroup-grouptype-property-outlook.md)|
|[Name](navigationgroup-name-property-outlook.md)|
|[NavigationFolders](navigationgroup-navigationfolders-property-outlook.md)|
|[Parent](navigationgroup-parent-property-outlook.md)|
|[Position](navigationgroup-position-property-outlook.md)|
|[Session](navigationgroup-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
