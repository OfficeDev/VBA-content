---
title: ContactsModule Object (Outlook)
keywords: vbaol11.chm3195
f1_keywords:
- vbaol11.chm3195
ms.prod: outlook
api_name:
- Outlook.ContactsModule
ms.assetid: fb183bd5-c72f-b38f-97e3-209a2a463d24
ms.date: 06/08/2017
---


# ContactsModule Object (Outlook)

Represents the  **Contacts** navigation module in the Navigation Pane of an explorer.


## Remarks

The  **ContactsModule** object, derived from the **[NavigationModule](navigationmodule-object-outlook.md)** object, provides access to the navigation groups contained in the **Contacts** navigation module of the Navigation Pane for an explorer. Use the **[GetNavigationModule](navigationmodules-getnavigationmodule-method-outlook.md)** method or the **[Item](navigationmodules-item-method-outlook.md)** method of the **[Modules](navigationpane-modules-property-outlook.md)** collection for the parent **[NavigationPane](navigationpane-object-outlook.md)** object to retrieve a **NavigationModule** object, then use the **[NavigationModuleType](navigationmodule-navigationmoduletype-property-outlook.md)** property of the **NavigationModule** object to retrieve the navigation module type. If the **NavigationModuleType** property is set to **olModuleContacts**, you can then cast the **NavigationModule** object reference as a **ContactsModule** object to access the **[NavigationGroups](contactsmodule-navigationgroups-property-outlook.md)** property for that navigation module.

You can use the  **[Visible](contactsmodule-visible-property-outlook.md)** property to determine if the navigation module is visible and the **[Position](contactsmodule-position-property-outlook.md)** property to return or set the display position of the navigation module within the Navigation Pane. You can use the **[Name](contactsmodule-name-property-outlook.md)** property to return the display name of the **Contacts** navigation module within the Navigation Pane.


## Properties



|**Name**|
|:-----|
|[Application](contactsmodule-application-property-outlook.md)|
|[Class](contactsmodule-class-property-outlook.md)|
|[Name](contactsmodule-name-property-outlook.md)|
|[NavigationGroups](contactsmodule-navigationgroups-property-outlook.md)|
|[NavigationModuleType](contactsmodule-navigationmoduletype-property-outlook.md)|
|[Parent](contactsmodule-parent-property-outlook.md)|
|[Position](contactsmodule-position-property-outlook.md)|
|[Session](contactsmodule-session-property-outlook.md)|
|[Visible](contactsmodule-visible-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
