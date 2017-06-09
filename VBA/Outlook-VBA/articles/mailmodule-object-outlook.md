---
title: MailModule Object (Outlook)
keywords: vbaol11.chm3193
f1_keywords:
- vbaol11.chm3193
ms.prod: outlook
api_name:
- Outlook.MailModule
ms.assetid: df20efe5-be5c-952d-c6b7-20c20a83fda0
ms.date: 06/08/2017
---


# MailModule Object (Outlook)

Represents the  **Mail** navigation module in the Navigation Pane of an explorer.


## Remarks

The  **MailModule** object, derived from the **[NavigationModule](navigationmodule-object-outlook.md)** object, provides read-only access to the navigation groups contained in the **Mail** navigation module of the Navigation Pane for an explorer. Use the **[GetNavigationModule](navigationmodules-getnavigationmodule-method-outlook.md)** method or the **[Item](navigationmodules-item-method-outlook.md)** method of the **[Modules](navigationpane-modules-property-outlook.md)** collection for the parent **[NavigationPane](navigationpane-object-outlook.md)** object to retrieve a **NavigationModule** object, then use the **[NavigationModuleType](navigationmodule-navigationmoduletype-property-outlook.md)** property of the **NavigationModule** object to retrieve the navigation module type. If the **NavigationModuleType** property is set to **olModuleMail**, you can then cast the **NavigationModule** object reference as a **MailModule** object to access the **[NavigationGroups](mailmodule-navigationgroups-property-outlook.md)** property for that navigation module.


 **Note**  Unlike other navigation modules, such as the  **[CalendarModule](calendarmodule-object-outlook.md)** object, you cannot create or delete navigation groups in the **MailModule** object.

You can use the  **[Visible](mailmodule-visible-property-outlook.md)** property to determine if the navigation module is visible, and the **[Position](mailmodule-position-property-outlook.md)** property to return or set the display position of the navigation module within the Navigation Pane. You can use the **[Name](mailmodule-name-property-outlook.md)** property to return the display name of the **Mail** navigation module within the Navigation Pane.


## Properties



|**Name**|
|:-----|
|[Application](mailmodule-application-property-outlook.md)|
|[Class](mailmodule-class-property-outlook.md)|
|[Name](mailmodule-name-property-outlook.md)|
|[NavigationGroups](mailmodule-navigationgroups-property-outlook.md)|
|[NavigationModuleType](mailmodule-navigationmoduletype-property-outlook.md)|
|[Parent](mailmodule-parent-property-outlook.md)|
|[Position](mailmodule-position-property-outlook.md)|
|[Session](mailmodule-session-property-outlook.md)|
|[Visible](mailmodule-visible-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
